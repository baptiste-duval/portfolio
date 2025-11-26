
function normalizeTextForComparison(s) {
    if (!s) return "";
    // Supprimer les accents
    return s.toString()
        .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        // Remplacer les caractères spéciaux par des espaces
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();
}

function verifierVMSuccess(message) {
    // Normalise le corps du mail en retirant tous les espaces et retours à la ligne
    const normalizedBody = message.getPlainBody().replace(/[\r\n\s]+/g, '').trim();
    // Regex pour trouver "Totalof X VMs" suivi de "VMsStatus Y Successful" (même avec des mots entre)
    const regex = /Totalof\s*(\d+)\s*VMs.*?VMsStatus\s*(\d+)\s*Successful/i;
    const match = normalizedBody.match(regex);
    if (match && match.length >= 3) {
        const totalVMs = parseInt(match[1], 10);
        const successfulVMs = parseInt(match[2], 10);
        if (totalVMs > 0 && totalVMs === successfulVMs) return "OK";
        return "NOK";
    }
    return null;
}

function getClientRefMatch(clientName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const feuilleObjets = ss.getSheetByName("Objets");
    if (!feuilleObjets || feuilleObjets.getLastRow() < 2) return null;

    const data = feuilleObjets.getRange(2, 1, feuilleObjets.getLastRow() - 1, 4).getValues();
    const clientNameNorm = normalizeTextForComparison(clientName);
    const clientNameUpper = clientName.toUpperCase(); 

    for (let i = 0; i < data.length; i++) {
        const [client, alias1, alias2] = data[i].map(d => d ? d.toString().trim() : "");
        const clientUpper = client.toUpperCase(); 
        
        // 1. Match Exact sur la Colonne A (client)
        if (normalizeTextForComparison(client) === clientNameNorm) return client;

        const alias1Norm = normalizeTextForComparison(alias1);
        const alias2Norm = normalizeTextForComparison(alias2);
        
        // 2. Match par Alias (Alias exact ou inclusion)
        if (alias1Norm.length > 3 && (alias1Norm === clientNameNorm || clientNameNorm.includes(alias1Norm))) return client;
        if (alias2Norm.length > 3 && (alias2Norm === clientNameNorm || clientNameNorm.includes(alias2Norm))) return client;
        
        // 3. Règle d'Inclusion pour Altaro/Nom court dans Nom long
        if (clientUpper.includes(clientNameUpper) && clientUpper.length > clientNameUpper.length) {
             // Évite les faux positifs trop courts ou non pertinents
             if (clientNameUpper.length > 5 || clientUpper.includes("ALTARO")) {
                return client;
             }
        }
    }
    return null;
}

function determineStatut(item) {
    const { message, ref, client } = item;
    const body = message.getPlainBody(); 
    
    // Détermination de la référence client pour les clients spéciaux qui n'ont pas de 'ref' lors du pré-traitement
    let currentRef = ref;
    if (!currentRef) {
        const clientRefName = getClientRefMatch(client);
        if (clientRefName) {
            const feuilleObjets = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Objets");
            const data = feuilleObjets.getRange(2, 1, feuilleObjets.getLastRow() - 1, 4).getValues();
            const refRow = data.find(row => row[0].toString().trim() === clientRefName);
            if (refRow) {
                currentRef = {
                    client: clientRefName,
                    alias2Exact: refRow[2] ? refRow[2].toString().trim() : "",
                    forceExact: (refRow[3] ? row[3].toString().trim() : "").indexOf('/') > -1
                };
            }
        }
    }
    
    // Si la référence n'est toujours pas trouvée, on suppose Inconnu, SAUF pour CloudAlly
    if (!currentRef && client !== "CLOUDALLY") return "Inconnu";
    
    const sujetTrim = message.getSubject().trim();
    const sujetNorm = normalizeTextForComparison(message.getSubject());

    // --- 1. RÈGLE SPÉCIALE CLOUDALLY (0 échec = OK) ---
    const isCloudAllyMail = client.toUpperCase().includes("CLOUDALLY") || message.getSubject().includes("CloudAlly Backup Summary");
    if (isCloudAllyMail) {
        // Recherche si 0 échec est mentionné
        if (message.getSubject().includes("0 Backup items Failed") || body.includes("0 Backup items Failed") || body.includes("0 items Failed")) {
            return "OK (CloudAlly 0 échec)";
        }
        // Si c'est un mail CloudAlly mais qu'il n'y a pas "0 échec", on le considère NOK par défaut
        return "NOK (CloudAlly échec)";
    }

    // --- 2. RÈGLES D'EXCLUSION (Échec Prioritaire RENFORCÉES) ---
    if (sujetNorm.includes("echec") || sujetNorm.includes("failed") || sujetNorm.includes("failure") ||
        sujetNorm.includes("missed its scheduled") || sujetNorm.includes("erreur") || 
        sujetNorm.includes("a manqué ses sauvegardes") || sujetNorm.includes("ont manqué leurs sauvegardes") ||
        (sujetNorm.includes("tache de sauvegarde") && sujetNorm.includes("echoue"))) {
        return "NOK";
    }

    // --- 3. RÈGLE SPÉCIALE SOCOPA (Rapports multiples/simples) ---
    if (client.toUpperCase().includes("SOCOPA")) {
        const rapportBlocks = body.split(/Rapport de la tâche/i).slice(1);
        let componentStatuses = [];
        rapportBlocks.forEach(block => {
            const statusMatch = block.match(/Statut:\s*(\w+)/i);
            const statut = statusMatch ? normalizeTextForComparison(statusMatch[1].trim()) : "ECHEC_INTERNE";
            const storageMatch = block.match(/Coffre de Stockage:\s*"(.*?)"/i);
            const storageName = storageMatch ? storageMatch[1].trim() : "Non Spécifié";
            componentStatuses.push({ storage: storageName, status: statut });
        });

        if (componentStatuses.length === 1) {
            const report = componentStatuses[0];
            // Vérification du statut détaillé si non "succes"
            const finalStatus = (report.status === "succes") ? "OK" : "NOK";
            const details = `${report.storage}: ${report.status}`;
            return `${finalStatus} (${details})`;
        }
        // Cas rare de rapports multiples dans un seul mail
        if (componentStatuses.length > 1) {
            const hasFailure = componentStatuses.some(r => r.status !== "succes");
            const finalStatus = hasFailure ? "NOK" : "OK";
            const details = componentStatuses.map(r => `${r.storage}: ${r.status}`).join(', ');
            return `${finalStatus} (${details})`;
        }
    }

    // --- 4. RÈGLES SPÉCIALES TOLÉRANTES ---
    if (client.toUpperCase().includes("R&D - ACTIVEBACKUP") && sujetNorm.includes("partiellement terminee")) return "OK";
    if (client.toUpperCase().includes("NOLLET - VEEAM") && sujetNorm.includes("minor warnings")) return "OK";
    

    // --- 5. RÈGLES D'ALTARO/HORNET (Intégration) ---
    if (message.getFrom().toLowerCase().includes("hornet") || message.getSubject().toLowerCase().includes("altaro")) {
        if (currentRef && currentRef.forceExact) {
            const possibles = currentRef.alias2Exact.split("|").map(s => s.trim());
            return possibles.some(p => sujetTrim === p) ? "OK" : "NOK (Alias2 Exact Altaro manquant)";
        }
    }
    
    // --- 6. RÈGLE STRICTE GÉNÉRALE (Force Exact) ---
    if (currentRef && currentRef.forceExact) {
        const possibles = currentRef.alias2Exact.split("|").map(s => s.trim());
        return possibles.some(p => sujetTrim === p) ? "OK" : "NOK (Alias2 Exact manquant)";
    }

    // --- 7. RÈGLE FLEXIBLE (Détection de Succès) ---
    const isSuccess = (
        sujetNorm.includes("succes") || sujetNorm.includes("succes sur") ||
        sujetNorm.includes("[success]") || sujetNorm.includes("ok") ||
        sujetNorm.includes("reussi") || sujetNorm.includes("complete successfully") || 
        sujetNorm.includes("termine avec succes") 
    );
    
    // Si aucun échec n'a été détecté avant, on utilise la règle flexible.
    return isSuccess ? "OK" : "NOK";
}

// =================================================================================
// FONCTIONS DE GESTION DU RAPPORT GROUPÉ (CONTROL PANEL / ALTARO FINAL)
// =================================================================================

/**
 * Parse le corps du mail de rapport groupé Control Panel/Altaro pour extraire le statut
 * de chaque client.
 * @param {GmailMessage} message Le message Gmail de rapport groupé.
 * @returns {Array<Object>} Liste des résultats pour chaque client.
 */
function parseGroupControlPanelReport(message) {
    const body = message.getPlainBody();
    
    // 1. Normalisation des séparateurs et découpage initial
    let cleanedBody = body.replace(/\|\s*\|\s*/g, ' ').replace(/\n\s*\|/g, '\n|');
    cleanedBody = cleanedBody.replace(/R&D;/, 'R&D'); 
    
    // Séparer à chaque nouvelle ligne de client (après un --- et avant un |)
    const clientBlocks = cleanedBody.split(/\n\s*---\s*\n\s*\|/); 
    
    const clientRefList = {};
    const exclusionKeywords = ["RAPPORT QUOTIDIEN", "CLICK HERE", "VMs", "VADE FRANCE", "DÉTAIL", "---", "AVENUE ANTOINE PINAY", "MACHINES VIRTUELLES", "HORS SITE", "REPRODUCTION", "VERIFICATION", "59510", "HEM, FRANCE", "NOV."]; 
    const totalVmsRegex = /Total de\s*(\d+)\s*VMs/i;
    
    // On ignore le premier bloc (qui est l'en-tête général du mail)
    clientBlocks.slice(1).forEach(block => {
        // Normaliser les lignes du bloc pour la recherche
        const lines = block.trim().split('\n').map(l => l.trim()).filter(l => l.length > 0);
        
        // --- 1. Détection du Nom du Client (Filtrage renforcé) ---
        let clientNameLine = lines.find(l => {
            const upCaseLine = l.toUpperCase();
            return l.length > 3 && 
                       !/^\d{2}\s+/.test(upCaseLine) && 
                       !exclusionKeywords.some(kw => upCaseLine.includes(kw.toUpperCase()));
        });
        
        if (!clientNameLine) return;

        let potentialClient = clientNameLine.replace(/^\|/, '').replace(/---$/, '').trim();
        potentialClient = potentialClient.replace(/Les alertes.*?\s*$/, '').trim();
        const currentClientName = potentialClient;
        
        if (clientRefList[currentClientName] || currentClientName.length < 3 || currentClientName.toLowerCase().includes("altaro")) return;
        
        // --- 2. Recherche des Totaux et Succès ---
        let totalVMs = 0;
        let successfulVMs = 0;
        
        const entireBlockText = block.replace(/[\r\n]/g, ' ').replace(/\s+/g, ' ').trim();
        const totalMatch = entireBlockText.match(totalVmsRegex);
        totalVMs = totalMatch ? parseInt(totalMatch[1], 10) : 0;
        
        const successLineRegex = /Succès\s*\|\s*(\d+)/i;
        const successLineMatch = block.match(successLineRegex);
        
        if (successLineMatch && successLineMatch.length > 1) {
            successfulVMs = parseInt(successLineMatch[1], 10);
        }
        
        if (totalVMs > 0 && successfulVMs === 0) {
            if (entireBlockText.includes("Échec") || entireBlockText.includes("Non Sauvegardé")) {
                 successfulVMs = -1; // Marquer comme échec connu
            }
        }


        // --- 3. Détermination du Statut ---
        const clientData = {
            client: currentClientName,
            total: totalVMs,
            succes: successfulVMs > 0 ? successfulVMs : 0,
            message: message,
            clientRefMatch: null 
        };
        
        // RÈGLE SPÉCIFIQUE AVRIL-BONDUES (2/3 = OK)
        if (normalizeTextForComparison(currentClientName).includes(normalizeTextForComparison("AVRIL-BONDUES")) && totalVMs === 3 && clientData.succes >= 2) {
            clientData.statut = "OK";
            clientData.commentaires = `ALTARO OK (Règle Avril-Bondues 2/3 : ${clientData.succes}/${totalVMs} VMs)`;
        } 
        // RÈGLE GÉNÉRALE
        else if (totalVMs === 0) {
            clientData.statut = "OK";
            clientData.commentaires = `ALTARO OK (0 VM à sauvegarder)`;
        } else if (clientData.succes === totalVMs && totalVMs > 0) {
            clientData.statut = "OK";
            clientData.commentaires = `ALTARO OK (${clientData.succes}/${totalVMs} VMs)`;
        } else {
            clientData.statut = "NOK";
            clientData.commentaires = `ALTARO ÉCHEC: ${clientData.succes}/${totalVMs} Succès`;
        }

        clientRefList[currentClientName] = clientData;
    });
    
    return Object.values(clientRefList).filter(item => item.client.length > 3);
}

/**
 * Log les rapports Control Panel dans une feuille dédiée et les marque comme lus.
 * @param {Array<GmailMessage>} messages Les messages Control Panel à logguer.
 */
function processControlPanelLogs(messages) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Control Panel Logs";
    let feuilleLog = ss.getSheetByName(sheetName);
    
    if (!feuilleLog) {
        feuilleLog = ss.insertSheet(sheetName);
        feuilleLog.appendRow(["Date", "Heure", "Expéditeur", "Sujet", "Corps du Mail"]);
        feuilleLog.setColumnWidth(5, 500);
    }
    
    messages.forEach(message => {
        const date = message.getDate();
        const dateString = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
        const timeString = Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm:ss");
        const body = message.getPlainBody() || message.getBody();
        feuilleLog.appendRow([dateString, timeString, message.getFrom(), message.getSubject(), body]);
        message.markRead();
    });
    
    if (feuilleLog.getLastRow() > 1) {
        const rangeToSort = feuilleLog.getRange(2, 1, feuilleLog.getLastRow() - 1, feuilleLog.getLastColumn());
        rangeToSort.sort({column: 1, ascending: false});
    }
}

/**
 * Trie les lignes Altaro selon une séquence stricte et inclut tous les éléments.
 *
 * @param {Array<string>} altaroLines Liste des lignes Altaro formatées (Client - ALTARO (x/y) = OK).
 * @returns {Array<string>} Liste triée.
 */
function sortAndFilterAltaroLines(altaroLines) {
    // Liste de tri alphabétique temporaire
    const altaroObjects = altaroLines.map(line => {
        const match = line.match(/^(.+?)\s+-\s+ALTARO/i);
        let clientName = match ? match[1].trim() : line;
        return { clientName, line };
    });

    // Tri alphabétique (pour la section Détails)
    const filteredAndSorted = altaroObjects
        .sort((a, b) => a.clientName.localeCompare(b.clientName))
        .map(obj => obj.line);

    // On retire les doublons si une même ligne détaillée Altaro est présente deux fois
    return [...new Set(filteredAndSorted)];
}

// =================================================================================
// FONCTIONS DE PRÉPARATION/ENVOI DE MAIL
// =================================================================================

/**
 * Crée un graphique circulaire représentant la répartition OK/NOK/Manquant.
 * @param {Array<string>} formulaireList La liste des statuts finaux par client.
 * @returns {Blob} L'image du graphique.
 */
function createPieChartBlob(formulaireList) {
    var okCount = 0;
    var nokCount = 0;
    var manquantCount = 0;
    
    formulaireList.forEach(function(item) {
        if (item.includes("= OK")) okCount++;
        else if (item.includes("= NOK")) nokCount++; 
        else if (item.includes("= Manquant")) manquantCount++; 
    });
    
    var dataTable = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, 'Statut')
        .addColumn(Charts.ColumnType.NUMBER, 'Nombre')
        .addRow(['✅ OK (' + okCount + ')', okCount])
        .addRow(['❌ NOK (' + nokCount + ')', nokCount]) 
        .addRow(['⚠️ Manquant (' + manquantCount + ')', manquantCount]) 
        .build();

    var chart = Charts.newPieChart()
        .setTitle('Répartition des Sauvegardes (OK vs NOK/Manquant) - Clients Référencés') 
        .setDataTable(dataTable)
        .setDimensions(450, 250) 
        .set3D()
        .setColors(['#1D8348', '#C0392B', '#F39C12']) // Vert, Rouge, Orange
        .build();
    return chart.getBlob().setName("pieChart.png");
}

/**
 * Envoie le rapport final par email.
 * @param {Array<string>} okList Liste des mails OK trouvés.
 * @param {Array<string>} nokList Liste des mails NOK trouvés.
 * @param {Array<string>} inconnuList Liste des mails de clients inconnus.
 * @param {Array<string>} formulaireList Statut final de tous les clients référencés.
 */
function envoyerRapportSimple(okList, nokList, inconnuList, formulaireList) {
    // Dédupliquer et trier
    okList = [...new Set(okList)];
    nokList = [...new Set(nokList)];
    inconnuList = [...new Set(inconnuList)];
    
    // Trier la liste du Statut Final: NOK (3), Manquant (2), OK (1), puis Alphabétique
    formulaireList.sort(function(a, b) {
        var aStatut = a.includes("= NOK") ? 3 : (a.includes("= Manquant") ? 2 : 1);
        var bStatut = b.includes("= NOK") ? 3 : (b.includes("= Manquant") ? 2 : 1);

        if (aStatut !== bStatut) return bStatut - aStatut; // 3 > 2 > 1
        return a.localeCompare(b); 
    });

    // Préparation de l'image
    var chartBlob, inlineImages = {};
    try {
        chartBlob = createPieChartBlob(formulaireList);
        inlineImages = { chartImage: chartBlob };
    } catch (e) {
        Logger.log("ATTENTION: Impossible de créer le graphique. Erreur: " + e.message);
    }

    // Création du HTML pour le statut final
    var formulaireHTML = formulaireList.map(function(x) {
        var color = x.includes("= NOK") ? '#C0392B' :
            x.includes("= Manquant") ? '#F39C12' : // Couleur Orange pour Manquant
            x.includes("= OK") ? '#1D8348' : '#2E86C1';
        return '<li style="color:' + color + '">' + x + '</li>';
    }).join('');
    
    // Construction du corps HTML du mail
    var corpsHTML = `
        <div style="font-family: Arial, sans-serif; font-size: 14px;">
            <h2 style="color:#2E86C1;">📊 Rapport Simple Sauvegardes Quotidien</h2>
            <p>Ce rapport vérifie la présence et le statut des sauvegardes (période : ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "EEE dd/MM/yyyy HH:mm:ss Z")}).</p>

            ${inlineImages.chartImage ? `
            <div style="text-align: center; margin: 20px 0;">
                <img src="cid:chartImage" alt="Diagramme Circulaire OK/NOK" style="max-width: 100%; height: auto; border: 1px solid #ddd; border-radius: 5px;">
            </div>
            ` : ''}
            <hr style="border: 0; height: 1px; background: #ddd;">

            <h3 style="color:#2E86C1;">🩵 Clients Référencés (Statut Final)</h3>
            <p>Statut des clients référencés : **OK**, **NOK**, ou **Manquant** (mail non trouvé).</p>
            <ul style="color:black; list-style-type: none; padding-left: 0;">
                ${formulaireHTML}
            </ul>
        `;
        
    corpsHTML += `<hr style="border: 0; height: 1px; background: #ddd;">`;
    
    if (nokList.length) {
        corpsHTML += `
            <h3 style="color:#C0392B;">❌ Détails des Échecs (NOK)</h3>
            <ul style="color:#C0392B;">
                ${nokList.map(x => `<li>${x}</li>`).join("")}
            </ul>
        `;
    }
    
    if (inconnuList.length) {
        corpsHTML += `
            <h3 style="color:#F39C12;">⚠️ Mails non Référencés (Clients Inconnus / Manquants Globaux)</h3>
            <ul style="color:#F39C12;">
                ${inconnuList.map(x => `<li>${x}</li>`).join("")}
            </ul>
        `;
    }
    
    if (okList.length) {
        corpsHTML += `
            <h3 style="color:#1D8348;">✅ Détails des Mails OK</h3>
            <ul style="color:#1D8348;">
                ${okList.map(x => `<li>${x}</li>`).join("")}
            </ul>
        `;
    }
    corpsHTML += "</div>";

    // Envoi du mail
    MailApp.sendEmail({
        to: "baptisteduval.bd62@gmail.com", // ⚠️ REMPLACER PAR VOTRE ADRESSE EMAIL
        subject: "📊 Rapport Simple Sauvegardes Quotidien",
        htmlBody: corpsHTML,
        inlineImages: inlineImages 
    });
    Logger.log("INFO: Le rapport a été envoyé par email.");
}

function genererRapportSauvegardeSimple() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let feuilleRapport = ss.getSheetByName("Rapport");
    if (!feuilleRapport) feuilleRapport = ss.insertSheet("Rapport");
    const header = ["Date", "Client", "Expéditeur", "Objet", "Statut", "VM Total", "Succès", "Détails"];
    
    const existingHeaderRange = feuilleRapport.getRange(1, 1, 1, Math.max(8, feuilleRapport.getLastColumn()));
    const existingHeader = existingHeaderRange.getValues()[0];
    if (existingHeader.join(",") !== header.join(",")) {
        feuilleRapport.getRange(1, 1, 1, header.length).setValues([header]);
    }

    const now = new Date();
    const start = new Date(now);
    const dayOfWeek = now.getDay(); // 0 = Dimanche, 1 = Lundi, ..., 6 = Samedi
    
    // --- LOGIQUE DE DATE WEEKEND (LUNDI) ---
    if (dayOfWeek === 1) { // Si c'est Lundi
        start.setDate(start.getDate() - 3); // Recule de 3 jours (Vendredi)
        Logger.log("Lundi détecté : Recherche depuis Vendredi 17h00.");
    } else {
        start.setDate(start.getDate() - 1); // Recule de 1 jour (Hier)
    }
    
    start.setHours(17, 0, 0, 0); 
    const query = 'label:backup_check after:' + Math.floor(start.getTime() / 1000);
    const threads = GmailApp.search(query);
    Logger.log(`Fenêtre d'analyse : depuis ${start.toLocaleString()} jusqu'à maintenant. ${threads.length} threads trouvés.`);

    // Liste des clients Altaro autorisés pour l'intégration automatique dans le statut final
    const authorizedAltaroClients = [
        "ARTOIS-FLEXIBLES", "ATS Groupe", "AVRIL-BONDUES", "CLEANINGBIO", "DT Froid", 
        "ETIQ-LEERS", "LYS-SANTE", "MBT Agencements", "MDZ Expertise & Conseils", 
        "MJL-PYCKAERT", "NCS-STAMAND", "OMEO-OIGNIES", "R&D", "TERRATECK-LESTREM", 
        "UTB", "VIVALANGUES", "VSG2C"
    ].map(name => normalizeTextForComparison(name)); // Normalisation pour la vérification

    // Préparation des références clients
    const feuilleObjets = ss.getSheetByName("Objets");
    const objetsRef = [];
    const formulaireStatus = {}; // Statut final (Manquant, OK, NOK)
    if (feuilleObjets && feuilleObjets.getLastRow() >= 2) {
        feuilleObjets.getRange(2, 1, feuilleObjets.getLastRow() - 1, 4).getValues().forEach(row => {
            const clientName = row[0] ? row[0].toString().trim() : "";
            const alias2Value = row[2] ? row[2].toString().trim() : "";
            const regleD = row[3] ? row[3].toString().trim() : "";
            if (clientName) {
                objetsRef.push({
                    client: clientName,
                    alias1Norm: normalizeTextForComparison(row[1] ? row[1].toString().trim() : ""),
                    alias2Norm: normalizeTextForComparison(alias2Value),
                    alias2Exact: alias2Value,
                    forceExact: (regleD.indexOf('/') > -1)
                });
                
                // Initialisation des statuts (Manquant par défaut)
                if (!formulaireStatus[clientName]) formulaireStatus[clientName] = "Manquant";
            }
        });
    }
    if (!formulaireStatus["1001 LOISIRS"]) {
        formulaireStatus["1001 LOISIRS"] = "Manquant";
    }


    const okList = [];
    const nokList = [];
    const inconnuList = [];
    let messagesToProcess = []; // Mails standards
    const processedMessageIds = new Set();
    const synergicMessages = [];
    const delangueMessages = [];
    const tostainMessages = []; // Liste pour les mails Tostain
    const cloudAllyMessages = []; // Liste pour les mails CloudAlly
    let controlPanelGroupResults = [];
    
    const nonMatchedAltaroClients = {}; 
    
    // --- 1. PARCOURS ET CLASSIFICATION DES THREADS ---
    threads.forEach(thread => {
        const allMessages = thread.getMessages();
        const messagesToAnalyze = allMessages.filter(msg => msg.getDate() >= start && msg.getDate() <= now);
        if (messagesToAnalyze.length === 0) return;
        
        // Traitement des rapports Control Panel (groupé)
        const controlPanelMessage = messagesToAnalyze.find(msg => msg.getSubject().includes("Rapport quotidien Control Panel"));
        if (controlPanelMessage && !processedMessageIds.has(controlPanelMessage.getId())) {
            processControlPanelLogs([controlPanelMessage]);
            const groupResults = parseGroupControlPanelReport(controlPanelMessage);
            controlPanelGroupResults = controlPanelGroupResults.concat(groupResults);
            processedMessageIds.add(controlPanelMessage.getId());
            return;
        }

        // Détermination du client le plus probable
        messagesToAnalyze.forEach(message => { 
            if (processedMessageIds.has(message.getId())) return;
            
            let bestMatch = { client: "Inconnu", ref: null, score: 0 };
            const sujetNorm = normalizeTextForComparison(message.getSubject() || "");
            const sujetUpper = message.getSubject().toUpperCase();
            
            let currentClient = "Inconnu";
            
            // Logique de classification précoce (clients spéciaux)
            if (sujetUpper.includes("CLOUDALLY")) {
                currentClient = "CLOUDALLY";
            } else if (sujetUpper.includes("DELANGUE")) {
                if (sujetUpper.includes("VEEAM") || sujetUpper.includes("JOB VEEAM DELANGUE")) {
                    currentClient = "DELANGUE - VEEAM";
                } else if (sujetUpper.includes("WASABI")) {
                    currentClient = "DELANGUE - WASABI";
                }
            } else if (sujetUpper.includes("SYNERGIC")) {
                 if (sujetUpper.includes("VEEAM") || sujetUpper.includes("JOB SYNERGIC") || sujetUpper.includes("SYNERGIC JOB")) {
                    currentClient = "SYNERGIC - VEEAM"; 
                } else if (sujetUpper.includes("WASABI")) {
                    currentClient = "SYNERGIC - WASABI";
                }
            } else if (sujetUpper.includes("TOSTAIN") && (sujetUpper.includes("RSYNC") || sujetUpper.includes("WASABI"))) {
                 if (sujetUpper.includes("RSYNC")) {
                    currentClient = "TOSTAIN - RSYNC"; 
                } else if (sujetUpper.includes("WASABI")) {
                    currentClient = "TOSTAIN - WASABI";
                }
            }
            
            if (currentClient === "Inconnu") {
                 // Recherche par référence client standard
                objetsRef.forEach(ref => {
                    let currentScore = 0;
                    if (ref.alias1Norm) {
                        ref.alias1Norm.split("|").forEach(a => { if (sujetNorm.indexOf(a) > -1) currentScore = a.length + 1000; });
                    } else if (ref.alias2Norm && sujetNorm.indexOf(ref.alias2Norm) > -1) {
                        currentScore = ref.alias2Norm.length;
                    }
                    if (sujetNorm.includes(normalizeTextForComparison(ref.client)) && sujetNorm.length > normalizeTextForComparison(ref.client).length) {
                        currentScore = Math.max(currentScore, normalizeTextForComparison(ref.client).length + 100);
                    }
                    
                    if (currentScore > bestMatch.score) {
                        bestMatch = { client: ref.client, ref: ref, score: currentScore };
                        currentClient = ref.client;
                    }
                });
            }
            
            bestMatch.client = currentClient;
            bestMatch.ref = objetsRef.find(r => r.client === bestMatch.client) || null;
            
            const item = { client: bestMatch.client, ref: bestMatch.ref, message: message };
            
            // Stockage dans les listes spécifiques
            if (bestMatch.client.includes("SYNERGIC")) {
                synergicMessages.push(item);
            } else if (bestMatch.client.includes("DELANGUE")) {
                delangueMessages.push(item);
            } else if (bestMatch.client.includes("TOSTAIN")) {
                tostainMessages.push(item);
            } else if (bestMatch.client === "CLOUDALLY") {
                cloudAllyMessages.push(item);
            } else {
                messagesToProcess.push(item);
            }
            processedMessageIds.add(message.getId());
        });
    });

    // --- 2. TRAITEMENT DES RÉSULTATS CONTROL PANEL (ALTARO) ---
    const allAltaroLines = [];
    controlPanelGroupResults.forEach(item => {
        const { client: clientTrouve, statut, total, succes, commentaires, message } = item;
        const statutSimple = statut.startsWith("OK") ? "OK" : "NOK";
        const clientNorm = normalizeTextForComparison(clientTrouve);
        
        let clientEstReference = false;
        let clientRefMatch = getClientRefMatch(clientTrouve);

        if (clientRefMatch && formulaireStatus[clientRefMatch]) {
            // Logique de consolidation weekend : OK écrase NOK/Manquant
            if (formulaireStatus[clientRefMatch] === "Manquant" || statutSimple === "OK") {
                formulaireStatus[clientRefMatch] = statutSimple;
            } 
            clientEstReference = true;
        } 
        
        // Logique de consolidation weekend pour les Altaro non référencés mais autorisés
        if (!clientEstReference && authorizedAltaroClients.includes(clientNorm)) {
             const clientNameForDisplay = `${clientTrouve} - ALTARO`;
             // Si le client n'est pas encore dans la liste ou si le nouveau statut est OK
             if (!nonMatchedAltaroClients[clientNameForDisplay] || statutSimple === "OK") {
                 nonMatchedAltaroClients[clientNameForDisplay] = statutSimple;
             }
             clientEstReference = true;
        }
        
        // Log dans la feuille de rapport
        const finalClientNameForSheet = clientEstReference ? clientRefMatch || `${clientTrouve} - ALTARO` : `Inconnu Altaro (${clientTrouve})`;
        feuilleRapport.appendRow([
            Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
            finalClientNameForSheet,
            message.getFrom(),
            "Rapport quotidien Control Panel",
            statutSimple, 
            total,
            succes,
            commentaires
        ]);

        if (clientEstReference) {
            const concisLine = `${clientTrouve} - ALTARO (${succes}/${total}) = ${statutSimple}`;
            allAltaroLines.push(concisLine);
        }
    });
    
    // Ajout des lignes Altaro aux détails (Dédupliquées et triées)
    const sortedAltaroLines = sortAndFilterAltaroLines(allAltaroLines);
    sortedAltaroLines.forEach(line => {
        if (line.includes("= OK")) okList.push(line);
        else nokList.push(line);
    });

    // --- 3. TRAITEMENT ET AGRÉGATION DES MESSAGES STANDARDS (OK Prioritaire sauf exceptions) ---
    messagesToProcess.forEach(item => {
    const { client: clientTrouve, message } = item;
    const statutDetaille = determineStatut(item);
    const statutSimple = statutDetaille.startsWith("OK") ? "OK" : (statutDetaille.startsWith("NOK") ? "NOK" : "Inconnu");
    const detailsMatch = statutDetaille.match(/\((.*)\)/);
    const details = detailsMatch ? detailsMatch[1] : "";

    let finalClient = clientTrouve;
    if (clientTrouve === "Inconnu") {
        const clientMatch = getClientRefMatch(message.getSubject());
        if (clientMatch) finalClient = clientMatch;

        // RÈGLE D'URGENCE POUR 1001 LOISIRS
        else if (message.getSubject().toUpperCase().includes("1000 ET 1 LOISIRS")) finalClient = "1001 LOISIRS";
        
        // Règle DSA (à garder)
        else if (normalizeTextForComparison(message.getSubject()).includes("dsa-lille-siege") || normalizeTextForComparison(message.getSubject()).includes("dsa lille siege")) finalClient = "DSA";
    }

    // ********** DE NOUVEAU CRITIQUE **********
    // Si la feuille "Objets" n'est pas lue correctement, "1001 LOISIRS" n'est pas initialisé
    // Nous devons le forcer à exister si le mail est trouvé.

    if (finalClient === "1001 LOISIRS" && !formulaireStatus["1001 LOISIRS"]) {
        formulaireStatus["1001 LOISIRS"] = "Manquant"; // Initialise le client si l'identification a réussi mais qu'il manque
    }
    
    // LOGIQUE D'AGRÉGATION GÉNÉRALE (elle doit se produire ici)
    else if (finalClient !== "Inconnu" && formulaireStatus[finalClient]) {
        // ... votre logique OK écrase NOK/Manquant
    }
        
        // Log dans la feuille de rapport (chaque message)
        feuilleRapport.appendRow([
            Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
            finalClient,
            message.getFrom(),
            message.getSubject(),
            statutSimple, 
            verifierVMSuccess(message) === "OK" ? 1 : 0, 
            verifierVMSuccess(message) === "OK" ? 1 : 0,
            details
        ]);
        
        // LOGIQUE SPÉCIALE D'AGRÉGATION SOCOPA: NOK écrase OK
        if (finalClient.toUpperCase().includes("SOCOPA") && formulaireStatus[finalClient]) {
            const currentFinalStatus = formulaireStatus[finalClient];
            
            // Si un mail NOK est trouvé, le statut final devient NOK, et il ne change plus
            if (statutSimple === "NOK") {
                 formulaireStatus[finalClient] = "NOK"; 
            } 
            // Si le statut actuel est Manquant, on le met à jour avec le statut du mail (OK ou Inconnu)
            else if (currentFinalStatus === "Manquant") {
                 formulaireStatus[finalClient] = statutSimple; 
            }
        } 
        // LOGIQUE D'AGRÉGATION GÉNÉRALE (OK Prioritaire)
        else if (finalClient !== "Inconnu" && formulaireStatus[finalClient]) {
            const currentFinalStatus = formulaireStatus[finalClient];
            
            if (currentFinalStatus === "Manquant") {
                 formulaireStatus[finalClient] = statutSimple;
            } else if (currentFinalStatus === "NOK" && statutSimple === "OK") {
                 formulaireStatus[finalClient] = "OK";
            }
        }
        
        // Ajout à la liste de détails (tous les mails sont gardés)
        const ligneRapportEmail = `${finalClient} | Objet: ${message.getSubject()} = ${statutDetaille}`;
        if (statutSimple === "OK") okList.push(ligneRapportEmail);
        else if (statutSimple === "NOK") nokList.push(ligneRapportEmail);
        else inconnuList.push(ligneRapportEmail);
});
    
          // --- 4. TRAITEMENT SPÉCIAL CLOUDALLY ---
      // --- 4. TRAITEMENT SPÉCIAL CLOUDALLY ---
      let cloudAllyGlobalStatut = "Manquant";
      let lastCloudAllyMessage = null;
      let isCloudAllyOK = false;

      if (cloudAllyMessages.length > 0) {
          // Trier par date pour trouver le plus récent (pour le log)
          const sortedMessages = cloudAllyMessages.sort((a, b) => b.message.getDate().getTime() - a.message.getDate().getTime()); 
          lastCloudAllyMessage = sortedMessages[0]; // Le message le plus récent pour le log final

          // Vérifier tous les messages pour trouver au moins un succès
          for (const item of cloudAllyMessages) {
              const statutDetaille = determineStatut(item);
              if (statutDetaille.startsWith("OK")) {
                  isCloudAllyOK = true;
                  break; // Succès trouvé, pas besoin de chercher plus loin
              }
          }
          
          // Déterminer le statut final :
          if (isCloudAllyOK) {
              cloudAllyGlobalStatut = "OK";
          } else {
              // Si aucun succès n'est trouvé, le statut est celui du dernier message (qui sera NOK selon determineStatut)
              const statutDetaille = determineStatut(lastCloudAllyMessage); 
              cloudAllyGlobalStatut = statutDetaille.startsWith("NOK") ? "NOK" : "Manquant";
          }

          // On met à jour CLOUDALLY uniquement s'il est référencé
          if (formulaireStatus["CLOUDALLY"]) formulaireStatus["CLOUDALLY"] = cloudAllyGlobalStatut;

          // Log du dernier mail CloudAlly
          // On utilise le dernier message reçu, même s'il n'est pas celui qui a donné le statut OK.
          // Pour le détail du mail OK, on ajoute une ligne de synthèse.
          // (Ajustez les lignes de log et de détails pour utiliser `lastCloudAllyMessage` ici)
          // ... (Le code de log continue)
          
          // ... (Log dans la feuille de rapport)
          feuilleRapport.appendRow([
              Utilities.formatDate(lastCloudAllyMessage.message.getDate(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
              "CLOUDALLY",
              lastCloudAllyMessage.message.getFrom(),
              lastCloudAllyMessage.message.getSubject(),
              cloudAllyGlobalStatut, "", "",
              `Statut Global: ${cloudAllyGlobalStatut}. (Détail: ${determineStatut(lastCloudAllyMessage)})`
          ]);

          const rapportLigneGlobal = `CLOUDALLY | Objet: ${lastCloudAllyMessage.message.getSubject()} = ${cloudAllyGlobalStatut} (Agrégation)`;
          if (cloudAllyGlobalStatut.includes("OK")) okList.push(rapportLigneGlobal);
          else if (cloudAllyGlobalStatut.includes("NOK")) nokList.push(rapportLigneGlobal);
          else inconnuList.push(rapportLigneGlobal); 
      }
    
    // --- 5. TRAITEMENT SPÉCIAL TOSTAIN (Double Vérification avec Règle 1er/2ème) ---
    
    // 1. Regroupement et Tri (plus ancien d'abord)
    const tostainRsyncMessages = tostainMessages
        .filter(item => item.client.includes("RSYNC"))
        .sort((a, b) => a.message.getDate().getTime() - b.message.getDate().getTime()); // Tri ascendant (plus ancien d'abord)

    const tostainWasabiMessages = tostainMessages
        .filter(item => item.client.includes("WASABI"))
        .sort((a, b) => a.message.getDate().getTime() - b.message.getDate().getTime()); // Tri ascendant (plus ancien d'abord)
        
    let finalTostainStatut = "NOK";
    let statutRsync, statutWasabi;
    let detailsRsync = "", detailsWasabi = "";
    
    // Fonction de vérification personnalisée pour un composant Tostain
    const checkTostainComponent = (messages, componentName) => {
        if (messages.length === 0) {
            return { statut: "Manquant", detail: "Manquant (Aucun mail trouvé)" };
        }
        
        const firstMessageItem = messages[0];
        const statutFirst = determineStatut(firstMessageItem);
        const statutSimpleFirst = statutFirst.startsWith("OK") ? "OK" : "NOK";
        
        let mailLog = `${componentName} (1er mail, le ${Utilities.formatDate(firstMessageItem.message.getDate(), Session.getScriptTimeZone(), "dd/MM HH:mm")})`;

        if (statutSimpleFirst === "OK") {
            return { statut: "OK", detail: `${mailLog} : OK` };
        }

        // Le premier est NOK, on vérifie le deuxième
        if (messages.length >= 2) {
            const secondMessageItem = messages[1];
            const statutSecond = determineStatut(secondMessageItem);
            const statutSimpleSecond = statutSecond.startsWith("OK") ? "OK" : "NOK";
            
            const mailLog2 = `${componentName} (2ème mail, le ${Utilities.formatDate(secondMessageItem.message.getDate(), Session.getScriptTimeZone(), "dd/MM HH:mm")})`;

            if (statutSimpleSecond === "OK") {
                return { statut: "OK", detail: `${mailLog} : NOK, ${mailLog2} : OK (Rattrapage)` };
            }
            
            return { statut: "NOK", detail: `${mailLog} : NOK, ${mailLog2} : NOK` };
        }
        
        // Un seul mail trouvé et il est NOK
        return { statut: "NOK", detail: `${mailLog} : NOK (Pas de second mail)` };
    };

    // Application pour RSYNC
    const rsyncResult = checkTostainComponent(tostainRsyncMessages, "RSYNC");
    statutRsync = rsyncResult.statut;
    detailsRsync = rsyncResult.detail;

    // Application pour WASABI
    const wasabiResult = checkTostainComponent(tostainWasabiMessages, "WASABI");
    statutWasabi = wasabiResult.statut;
    detailsWasabi = wasabiResult.detail;
    
    // Logique globale finale 
    if (statutRsync === "OK" && statutWasabi === "OK") {
         finalTostainStatut = "OK";
    } else if (statutRsync === "Manquant" && statutWasabi === "Manquant") {
         finalTostainStatut = "Manquant";
    } else {
         finalTostainStatut = "NOK"; 
    }
    
    // Mise à jour formulaireStatus
    if (formulaireStatus["TOSTAIN - RSYNC"]) formulaireStatus["TOSTAIN - RSYNC"] = statutRsync;
    if (formulaireStatus["TOSTAIN - WASABI"]) formulaireStatus["TOSTAIN - WASABI"] = statutWasabi;

    // Log Global Tostain
    feuilleRapport.appendRow([
        Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
        "TOSTAIN (Global)",
        "Synthèse",
        "Double Vérification RSYNC + WASABI (Règle 1er/2ème)",
        finalTostainStatut, "", "",
        `RSYNC: ${detailsRsync}, WASABI: ${detailsWasabi}`
    ]);
    
    // Mise à jour des listes de détails (okList/nokList/inconnuList)
    const rapportLigneTostain = `TOSTAIN (RSYNC et WASABI) | Statut Global: ${finalTostainStatut} - Détails: RSYNC: ${detailsRsync} / WASABI: ${detailsWasabi}`;
    if (finalTostainStatut.includes("OK")) okList.push(rapportLigneTostain);
    else if (finalTostainStatut.includes("NOK")) nokList.push(rapportLigneTostain);
    else inconnuList.push(rapportLigneTostain); 


    // --- 6. TRAITEMENT SPÉCIAL SYNERGIC (Double Vérification Wasabi + Veeam) ---
    // NOTE: Pour Synergic, la logique reste simple (prend le premier trouvé pour chaque composant)
    const synergicVeeam = synergicMessages.find(item => item.client.includes("VEEAM"));
    const synergicWasabi = synergicMessages.find(item => item.client.includes("WASABI"));
    
    let statutSVeeam = synergicVeeam ? determineStatut(synergicVeeam) : "Manquant"; 
    let statutSWasabi = synergicWasabi ? determineStatut(synergicWasabi) : "Manquant"; 
    
    let finalSynergicStatut = "NOK";
    if (statutSVeeam.includes("OK") && statutSWasabi.includes("OK")) {
         finalSynergicStatut = "OK";
    } else if (statutSVeeam.includes("Manquant") && statutSWasabi.includes("Manquant")) {
         finalSynergicStatut = "Manquant";
    } else if (statutSVeeam.includes("Manquant") || statutSWasabi.includes("Manquant")) {
         finalSynergicStatut = "NOK"; 
    }
    
    if (formulaireStatus["SYNERGIC - VEEAM"]) formulaireStatus["SYNERGIC - VEEAM"] = statutSVeeam.includes("Manquant") ? "Manquant" : statutSVeeam.split(" ")[0];
    if (formulaireStatus["SYNERGIC - WASABI"]) formulaireStatus["SYNERGIC - WASABI"] = statutSWasabi.includes("Manquant") ? "Manquant" : statutSWasabi.split(" ")[0];
    
    // Log Global Synergic
    feuilleRapport.appendRow([
        Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
        "SYNERGIC (Global)",
        "Synthèse",
        "Double Vérification VEEAM + WASABI",
        finalSynergicStatut, "", "",
        `WASABI: ${statutSWasabi}, VEEAM: ${statutSVeeam}`
    ]);
    
    const rapportLigneGlobal = `SYNERGIC (WASABI et VEEAM) | Statut Global : ${finalSynergicStatut}`;
    if (finalSynergicStatut.includes("OK")) okList.push(rapportLigneGlobal);
    else if (finalSynergicStatut.includes("NOK")) nokList.push(rapportLigneGlobal);
    else inconnuList.push(rapportLigneGlobal); 


    const delangueVeeam = delangueMessages.find(item => item.client.includes("VEEAM"));
    const delangueWasabi = delangueMessages.find(item => item.client.includes("WASABI"));
    
    let statutDVeeam = delangueVeeam ? determineStatut(delangueVeeam) : "Manquant"; 
    let statutDWasabi = delangueWasabi ? determineStatut(delangueWasabi) : "Manquant"; 

    let finalDelangueStatut = "NOK";
    if (statutDVeeam.includes("OK") && statutDWasabi.includes("OK")) {
         finalDelangueStatut = "OK";
    } else if (statutDVeeam.includes("Manquant") && statutDWasabi.includes("Manquant")) {
         finalDelangueStatut = "Manquant";
    } else if (statutDVeeam.includes("Manquant") || statutDWasabi.includes("Manquant")) {
         finalDelangueStatut = "NOK"; 
    }

    if (formulaireStatus["DELANGUE - VEEAM"]) formulaireStatus["DELANGUE - VEEAM"] = statutDVeeam.includes("Manquant") ? "Manquant" : statutDVeeam.split(" ")[0];
    if (formulaireStatus["DELANGUE - WASABI"]) formulaireStatus["DELANGUE - WASABI"] = statutDWasabi.includes("Manquant") ? "Manquant" : statutDWasabi.split(" ")[0];
    
    // Log Global Delangue
    feuilleRapport.appendRow([
        Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
        "DELANGUE (Global)",
        "Synthèse",
        "Double Vérification VEEAM + WASABI",
        finalDelangueStatut, "", "",
        `WASABI: ${statutDWasabi}, VEEAM: ${statutDVeeam}`
    ]);
    
    const rapportLigneDelangue = `DELANGUE (WASABI et VEEAM) | Statut Global : ${finalDelangueStatut}`;
    if (finalDelangueStatut.includes("OK")) okList.push(rapportLigneDelangue);
    else if (finalDelangueStatut.includes("NOK")) nokList.push(rapportLigneDelangue);
    else inconnuList.push(rapportLigneDelangue); 

    Object.assign(formulaireStatus, nonMatchedAltaroClients);
    
    const formulaireList = Object.keys(formulaireStatus).map(client => {
        const statut = formulaireStatus[client];
        // FORMATAGE FINAL: Afficher uniquement = OK, = NOK, ou = Manquant
        return `${client} = ${statut}`;
    });

    envoyerRapportSimple(okList, nokList, inconnuList, formulaireList);
}