//==============================================================================
// GLOBAL VARIABLES & CONSTANTS
//==============================================================================

// --- Current User's Dynamic Sheet IDs (set upon app access) ---
let CURRENT_USER_PRODUSE_ID = null;
let CURRENT_USER_COMENZI_ID = null;
let CURRENT_USER_DATE_FIRMA_ID = null;
let CURRENT_USER_FACTURI_ID = null;

// --- Thresholds & Specific IDs ---
const LOW_STOCK_THRESHOLD = 20; // For KPI calculation
const OPENAI_CHAT_ENDPOINT = "https://api.openai.com/v1/chat/completions";
const OPENAI_MODEL = "gpt-3.5-turbo"; // Or "gpt-4o"
const DEFAULT_VAT_RATE = 0.21; // NEW: Asigură-te că această linie este prezentă și corectă


// --- Master Template & User Management Sheet IDs (REPLACE PLACEHOLDERS!) ---
// These are the IDs of the sheets/folders YOU manage as the admin/developer.
// They serve as templates or mapping tables for all users.
const MASTER_PRODUSE_TEMPLATE_ID = "1y6z-eK2CcXPpeeSMx_TiGMvTjP52FXaWP00dZaVQFqM";
const MASTER_COMENZI_TEMPLATE_ID = "10dxElNvJqAz0Ma8IY_c0-p7Z9UxXnubM2ZOIBsuHX3I";
const MASTER_DATE_FIRMA_TEMPLATE_ID = "1l-3V0WcK4pK1cKLDZv1t2HtIFCw7m3LoFcRLeVg4IEg"; // Replace with your actual ID
const MASTER_FACTURI_TEMPLATE_ID = "1gI3XjTzpskjj4QA-UfPV1OI773BLo4Xh0RVefzZfxuo";
const USER_DATA_ROOT_FOLDER_ID = "1TX3G8BD1VWwmOi7O29tGaVdnlxVqdP9L";        // The folder where each user's dedicated folder will be created.
const USER_MAPPING_SHEET_ID = "192uEShs8r9PXH4iDdKwy2AlAVUGDkODrZEZAx0ekNAc";        // Sheet mapping user emails to their dedicated sheet IDs.
const AUTHORIZED_USERS_SHEET_ID = "1iewcwPqJWZWsV_HDx2WXPQxer3sfPPWvS9ukqucL_GA"; // Sheet listing authorized user emails.

// --- Invoice Template & Folder IDs ---
const INVOICE_TEMPLATE_ID = "13Ag4xJM9zzUx4ukjUfkCYpLf3U2H3nmUy_Jq0UT3qgY"; // E.g., "10r7wd8pN3zO_S-Rbwt4wM76WmxlPiOMfV_1XJnUMRqA"


//==============================================================================
// CORE APPS SCRIPT UTILITIES & WEB APP ENTRY POINT
//==============================================================================

/**
 * Main Web App Entry Point.
 * This function runs when the Web App URL is accessed via a GET request.
 * It serves the HTML content.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Orderly AI Chat');
}

/**
 * Logs an error message originating from the frontend (client-side).
 * This function is called by the frontend's withFailureHandler for debugging.
 * @param {string} errorMessage The error message from the client.
 */
function logServerError(errorMessage) {
  Logger.log(`Frontend Error: ${errorMessage}`);
}


//==============================================================================
// USER MANAGEMENT & ONBOARDING FUNCTIONS
//==============================================================================

/**
 * Checks if a user's email is present in the Authorized_Users sheet and retrieves their username.
 * @param {string} userEmail The email to check for authorization.
 * @returns {string|boolean} The user's username if authorized, or false otherwise.
 */
function checkUserAuthorization(userEmail) {
  try {
    const ss = SpreadsheetApp.openById(AUTHORIZED_USERS_SHEET_ID);
    const sheet = ss.getSheets()[0]; // Assuming authorized users are on the first tab
    const data = sheet.getDataRange().getValues(); // Get all data

    // Assuming headers are in row 1: User Email (index 0), User Name (index 1)
    const emailColumnIndex = data[0].indexOf("User Email");
    const nameColumnIndex = data[0].indexOf("User Name");

    if (emailColumnIndex === -1 || nameColumnIndex === -1) {
      Logger.log("Error: 'User Email' or 'User Name' column not found in Authorized_Users sheet headers.");
      return false; // Cannot authorize if columns are missing
    }

    // Iterate through rows to find the user
    for (let i = 1; i < data.length; i++) { // Skip header row
      if (data[i][emailColumnIndex].toString().toLowerCase() === userEmail.toLowerCase()) {
        const userName = data[i][nameColumnIndex]; // Get the username
        Logger.log(`Authorization granted for ${userEmail}. Username: ${userName}`);
        return userName; // Return the username if found
      }
    }

    Logger.log(`Access denied: ${userEmail} not found in Authorized_Users sheet.`);
    return false; // Not found
  } catch (e) {
    Logger.log(`Error checking user authorization: ${e.message}`);
    return false; // Error occurred
  }
}

/**
 * Gets or creates the dedicated Google Sheet IDs for the current user.
 * If sheets don't exist, it copies them from master templates and records in mapping sheet.
 * @param {string} userEmail The email of the current user.
 * @returns {Object|null} An object {folderId, produseId, comenziId, dateFirmaId, facturiId} or null if error.
 */
function getUserSheets(userEmail) {
  try {
    const mappingSs = SpreadsheetApp.openById(USER_MAPPING_SHEET_ID);
    const mappingSheet = mappingSs.getSheets()[0];
    const data = mappingSheet.getDataRange().getValues();

    let userRow = -1;
    let userSheetIds = null;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userEmail) {
        userRow = i;
        userSheetIds = {
          folderId: data[i][1],
          produseId: data[i][2],
          comenziId: data[i][3],
          dateFirmaId: data[i][4],
          facturiId: data[i][5]
        };
        Logger.log(`Found existing sheets for user: ${userEmail}`);
        break;
      }
    }

    if (userRow === -1) {
      Logger.log(`Creating new sheets for user: ${userEmail}`);
      const userDataRootFolder = DriveApp.getFolderById(USER_DATA_ROOT_FOLDER_ID);
      const userFolderName = `Orderly Data - ${userEmail}`;
      const userFolder = userDataRootFolder.createFolder(userFolderName);

      const masterProduse = DriveApp.getFileById(MASTER_PRODUSE_TEMPLATE_ID);
      const masterComenzi = DriveApp.getFileById(MASTER_COMENZI_TEMPLATE_ID);
      const masterDateFirma = DriveApp.getFileById(MASTER_DATE_FIRMA_TEMPLATE_ID);
      const masterFacturi = DriveApp.getFileById(MASTER_FACTURI_TEMPLATE_ID);

      const userProduseFile = masterProduse.makeCopy(`Produse - ${userEmail}`, userFolder);
      const userComenziFile = masterComenzi.makeCopy(`Comenzi - ${userEmail}`, userFolder);
      const userDateFirmaFile = masterDateFirma.makeCopy(`Date firma - ${userEmail}`, userFolder);
      const userFacturiFile = masterFacturi.makeCopy(`Facturi - ${userEmail}`, userFolder);

      userSheetIds = {
        folderId: userFolder.getId(),
        produseId: userProduseFile.getId(),
        comenziId: userComenziFile.getId(),
        dateFirmaId: userDateFirmaFile.getId(),
        facturiId: userFacturiFile.getId()
      };

      mappingSheet.appendRow([
        userEmail,
        userSheetIds.folderId,
        userSheetIds.produseId,
        userSheetIds.comenziId,
        userSheetIds.dateFirmaId,
        userSheetIds.facturiId
      ]);
      Logger.log(`Successfully created and mapped sheets for ${userEmail}.`);
    }
    return userSheetIds;
  } catch (e) {
    Logger.log(`Error in getUserSheets: ${e.message}`);
    return null;
  }
}

/**
 * Retrieves the URLs for the current user's dedicated sheets.
 * Called by frontend via google.script.run.
 * @param {Object} sheetIds An object containing {produseId, comenziId, dateFirmaId, facturiId} for the current user.
 * @returns {Object|null} An object {produseUrl, comenziUrl, dateFirmaUrl, facturiUrl} or null if IDs are not set.
 */
function getUserSheetUrls(sheetIds) {
  if (!sheetIds || !sheetIds.produseId || !sheetIds.comenziId || !sheetIds.dateFirmaId || !sheetIds.facturiId) {
    Logger.log("Error: User sheet IDs object is incomplete or not provided. Cannot get URLs.");
    return null;
  }
  return {
    produseUrl: `https://docs.google.com/spreadsheets/d/${sheetIds.produseId}/edit`,
    comenziUrl: `https://docs.google.com/spreadsheets/d/${sheetIds.comenziId}/edit`,
    dateFirmaUrl: `https://docs.google.com/spreadsheets/d/${sheetIds.dateFirmaId}/edit`,
    facturiUrl: `https://docs.google.com/spreadsheets/d/${sheetIds.facturiId}/edit`
  };
}

/**
 * Gets the email of the active user accessing the web app, checks authorization,
 * and sets their specific sheet IDs globally if authorized.
 * @returns {Object|null} An object {userEmail, userName, sheetIds: {...}} if authorized, or null if unauthorized/error.
 */
function getUserEmail() {
  const userEmail = Session.getActiveUser().getEmail() || "";
  Logger.log(`Web App accessed by (raw): ${userEmail}`);

  if (userEmail === "") {
    Logger.log("User email is empty/unresolved. Cannot check authorization.");
    return null;
  }

  const userName = checkUserAuthorization(userEmail); // CHANGED: checkUserAuthorization now returns username or false

  if (!userName) { // If userName is false (meaning not authorized)
    Logger.log(`Access denied for unauthorized user: ${userEmail}`);
    return null;
  }

  const userSheetIds = getUserSheets(userEmail);

  if (userSheetIds) {
    // Set global variables for the current user's sheet IDs
    CURRENT_USER_PRODUSE_ID = userSheetIds.produseId;
    CURRENT_USER_CLIENTI_ID = userSheetIds.clientiId; // This is still here even if Clienti sheet removed, just to ensure no error
    CURRENT_USER_COMENZI_ID = userSheetIds.comenziId;
    CURRENT_USER_DATE_FIRMA_ID = userSheetIds.dateFirmaId;

    Logger.log(`Set current user sheet IDs: Produse=${CURRENT_USER_PRODUSE_ID}, Comenzi=${CURRENT_USER_COMENZI_ID}, Date Firma=${CURRENT_USER_DATE_FIRMA_ID}`);

    // Return the full user object including email, username, and sheet IDs for frontend
    return { userEmail: userEmail, userName: userName, sheetIds: userSheetIds }; // CHANGED: Added userName
  } else {
    Logger.log(`Error: Could not get/create sheets for authorized user ${userEmail}.`);
    return null;
  }
}
/**
 * Obține sau creează ID-urile foilor dedicate utilizatorului curent și le setează global.
 * Aceasta funcție ar trebui apelată la începutul oricărei execuții backend care necesită acces la foile utilizatorului.
 * @returns {Object|null} Un obiect {produseId, comenziId, dateFirmaId, facturiId} sau null în caz de eroare/neautorizat.
 */
function getOrCreateAndSetUserSheetIds() {
  const userEmail = Session.getActiveUser().getEmail() || "";
  if (userEmail === "" || !checkUserAuthorization(userEmail)) {
    Logger.log("Utilizator neidentificat sau neautorizat. Nu se pot seta ID-urile foilor.");
    return null;
  }

  const userSheetIds = getUserSheets(userEmail);

  if (userSheetIds) {
    CURRENT_USER_PRODUSE_ID = userSheetIds.produseId;
    CURRENT_USER_COMENZI_ID = userSheetIds.comenziId;
    CURRENT_USER_DATE_FIRMA_ID = userSheetIds.dateFirmaId;
    CURRENT_USER_FACTURI_ID = userSheetIds.facturiId;
    Logger.log(`ID-uri foi setate global pentru execuție: Produse=${CURRENT_USER_PRODUSE_ID}, Comenzi=${CURRENT_USER_COMENZI_ID}, Date Firma=${CURRENT_USER_DATE_FIRMA_ID}, Facturi=${CURRENT_USER_FACTURI_ID}`);
    return userSheetIds;
  } else {
    Logger.log(`Eroare: Nu s-au putut obține/crea foile pentru utilizatorul ${userEmail}.`);
    return null;
  }
}


//==============================================================================
// CONVERSATION STATE MANAGEMENT
//==============================================================================

/**
 * Setează starea conversației pentru utilizatorul curent.
 * @param {string} state Numele stării (ex: "awaiting_order_confirmation").
 * @param {Object|null} data Datele asociate stării (ex: detalii comandă).
 */
function setConversationState(state, data) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('conversationState', state);
  if (data) {
    userProperties.setProperty('conversationData', JSON.stringify(data));
  } else {
    userProperties.deleteProperty('conversationData');
  }
  Logger.log(`Starea conversației setată la: ${state}`);
}

/**
 * Recuperează starea conversației pentru utilizatorul curent.
 * @returns {Object} Un obiect { state: string, data: Object|null }.
 */
function getConversationState() {
  const userProperties = PropertiesService.getUserProperties();
  const state = userProperties.getProperty('conversationState');
  const dataString = userProperties.getProperty('conversationData');
  const data = dataString ? JSON.parse(dataString) : null;
  return { state: state, data: data };
}

/**
 * Șterge starea conversației pentru utilizatorul curent.
 */
function clearConversationState() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('conversationState');
  userProperties.deleteProperty('conversationData');
  Logger.log("Starea conversației a fost ștearsă.");
}


/**
 * Gestionează fluxul de emitere a facturii după confirmarea inițială,
 * solicitând CUI/Email sau procesând comanda finală.
 * Aceasta este apelată de processUserMessage dacă detectează o stare de "awaiting_client_cui_and_email".
 * @param {string} userMessage Mesajul utilizatorului (care ar putea conține CUI/Email).
 * @param {Object} orderDetails Detaliile comenzii stocate în starea conversației.
 * @returns {Object} Un obiect { message: string, buttons: Array<string>|null }.
 */
function handleInvoiceIssuanceFlow(userMessage, orderDetails) {
  Logger.log(`Intrat în handleInvoiceIssuanceFlow cu mesaj: ${userMessage} și detalii comandă: ${JSON.stringify(orderDetails)}`);

  let finalResponse = "A apărut o problemă la procesarea facturii. Te rog încearcă din nou.";
  let responseButtons = null;

  // Extragem CUI-ul și email-ul din mesajul utilizatorului
  const userInputParts = userMessage.split(' ').map(s => s.trim());

  let clientCui = null;
  let clientEmail = null;

  for (const part of userInputParts) {
    if (part.includes('@') && part.includes('.')) {
      clientEmail = part;
    } else if (!isNaN(part) || part.toLowerCase().startsWith('ro')) {
      clientCui = part;
    }
  }

  // Verificăm dacă am extras cu succes CUI și Email
  if (clientCui && clientEmail) {
    Logger.log(`CUI/Email extrase: CUI=${clientCui}, Email=${clientEmail}`);
    finalResponse = handleOrderCreation(
      orderDetails.products,
      clientCui,
      clientEmail
    );
    clearConversationState(); // Șterge starea după procesare finală
  } else {
    finalResponse = "Nu am putut extrage CUI-ul și/sau adresa de email. Te rog, furnizează-le clar (ex: '47315510 test@example.com').";
    setConversationState("awaiting_client_cui_and_email", orderDetails);
  }

  return { message: finalResponse, buttons: responseButtons };
}


//==============================================================================
// CHATBOT CORE FUNCTIONS (OpenAI API Interaction & Prompt Management)
//==============================================================================

/**
 * Retrieves the OpenAI API key securely from script properties.
 * @returns {string} The OpenAI API key.
 */
function getOpenAIApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error("OpenAI API key not found in script properties. Please set 'OPENAI_API_KEY'.");
  }
  return apiKey;
}

/**
 * Sends a prompt to the ChatGPT API and returns the AI's response.
 * @param {string} userPrompt The user's message to send to ChatGPT.
 * @returns {string|null} The AI's response text, or null if an error occurs.
 */
function callChatGPT(userPrompt) {
  try {
    const apiKey = getOpenAIApiKey();

    const headers = {
      "Authorization": "Bearer " + apiKey,
      "Content-Type": "application/json"
    };

    const payload = {
      "model": OPENAI_MODEL,
      "messages": [
          {"role": "system", "content": `You are an AI inventory and order management assistant for a business in Romania.
        Your primary goal is to help users check product stock, process orders, and answer key performance indicator (KPI) questions.
        Users will interact with you in Romanian.

        **RESPOND ONLY WITH A JSON OBJECT IF YOU IDENTIFY AN INTENT.**
        **DO NOT ADD ANY OTHER TEXT OUTSIDE THE JSON for these intents.**

        --- Intents and JSON Formats ---

        1.  **Check Stock:**
            User query about stock. Identify 'product_id'.
            If 'product_id' is not provided by the user, omit it from the JSON.
            If the user mentions a product name that does NOT match any known product, use the "Unknown Product" intent instead.
            JSON: { "intent": "check_stock", "product_id": "P001" }

        2.  **Add Order:** // PROMPT REFINAT PENTRU FLEXIBILITATE ȘI PRODUSE MULTIPLE
            User wants to place an order. Identify 'products' (a list of product_id and quantity), 'customer_id' (CUI), and 'customer_email'.
            **'products' should be an array of objects: [{"product_id": "P001", "quantity": 5}, {"product_id": "P002", "quantity": 10}].**
            **Return ONLY the fields that are clearly provided by the user.**
            If some are missing, just omit them from the JSON. Do NOT use "missing_order_details" unless no order details at all.
            Example 1 (multiple products): { "intent": "add_order", "products": [{"product_id": "P001", "quantity": 5}, {"product_id": "P002", "quantity": 10}], "customer_id": "47315510", "customer_email": "client@email.com" }
            Example 2 (single product): { "intent": "add_order", "products": [{"product_id": "P001", "quantity": 10}], "customer_id": "47315510", "customer_email": "client@email.com" }
            Example 3 (product and quantity only): { "intent": "add_order", "products": [{"product_id": "P001", "quantity": 10}] }
            Example 4 (just product name): { "intent": "add_order", "products": [{"product_id": "P001"}] }
            Example 5 (just quantity): { "intent": "add_order", "products": [{"quantity": 5}] }

        3.  **Unknown Product:**
            If a user mentions a product name that you **CANNOT CONFIDENTLY MAP** to an **EXACT** ID_Produs from the provided list, use this intent.
            **DO NOT GUESS OR INFER A PRODUCT_ID IF IT'S NOT IN THE LIST.**
            JSON: { "intent": "unknown_product", "query": "original_user_product_name_here" }
            Example for unknown product: User asks "Cati trandafiri am in stoc?", you respond: { "intent": "unknown_product", "query": "trandafiri" }

        4.  **Get Total Products:**
            User asks for the total count of products in inventory.
            Examples: "Câte produse am pe stoc?", "Numărul total de produse.", "Total produse în stoc.", "Câte tipuri articole sunt în inventar?", "Raport produse totale."
            JSON: { "intent": "get_total_products" }

        5.  **Get Total Stock Value:**
            User asks for the total monetary value of all stock.
            Examples: "Cât valorează tot stocul?", "Valoarea totală a inventarului.", "Care este valoarea totală a stocului?"
            JSON: { "intent": "get_total_stock_value" }

        6.  **Get Number of Orders:**
            User explicitly asks for the *total count* of all orders placed or recorded. This is a summary metric, not about a specific order.
            Examples: "Câte comenzi am avut în total?", "Care este numărul total de comenzi?", "Raport număr comenzi."
            JSON: { "intent": "get_num_orders" }

        7.  **Get Total Order Revenue:**
            User asks for the total monetary value of all orders.
            Examples: "Care este venitul total din comenzi?", "Cât am încasat din vânzări?", "Venit total comenzi."
            JSON: { "intent": "get_total_order_revenue" }

        8.  **Get Low Stock Count:**
            User asks for the name of products with low stock. Assume "low stock" means less than 20 units.
            Examples: "Câte produse am cu stoc scăzut?", "Numărul de produse cu stoc redus.", "Produse puține în stoc."
            JSON: { "intent": "get_low_stock_count" }

        9.  **Missing Order Details:**
            If the user asks to add an order but does NOT provide ALL of product_id, quantity, customer_id, AND customer_email.
            JSON: { "intent": "missing_order_details", "message": "Pentru a procesa comanda, am nevoie de produs, cantitate, CUI și adresa de email. Ex: 'comanda 10 tricouri pentru 47315510 test@email.com'." }


           10. **Get Capabilities:** // NEW INTENT
            User asks what the AI can do or its functionalities.
            Examples: "Ce poți face?", "Care sunt funcționalitățile tale?", "Cu ce mă poți ajuta?"
            JSON: { "intent": "get_capabilities", "capabilities": ["gestionarea stocului", "gestionarea comenzilor", "facturare automată", "indicatori de performanță"] }
        --- Product Mapping (for use in product_id) ---

        Products:
        - Nume: "Tricou", ID_Produs: "P001"
        - Nume: "Bere", ID_Produs: "P002"


            --- General Questions ---
            If the user's query is general and does not match any of the above intents, respond naturally in Romanian. Do NOT use JSON for general questions.
            `},
         {"role": "user", "content": userPrompt}
      ],
      "max_tokens": 150
    };

    const options = {
      "method": "post",
      "headers": headers,
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    Logger.log(`Sending prompt to ChatGPT: ${userPrompt}`);
    const response = UrlFetchApp.fetch(OPENAI_CHAT_ENDPOINT, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);
      const assistantMessage = jsonResponse.choices[0].message.content.trim();
      Logger.log(`ChatGPT response: ${assistantMessage}`);
      return assistantMessage;
    } else {
      Logger.log(`ChatGPT API Error: Code ${responseCode}, Body: ${responseBody}`);
      return `Error from AI: ${responseBody}`;
    }
  } catch (e) {
    Logger.log(`Exception calling ChatGPT API: ${e.message}`);
    return `An internal error occurred: ${e.message}`;
  }
}



/**
 * Processes a user message, sends it to ChatGPT, interprets the intent,
 * interacts with Google Sheets, and returns the AI's response.
 * This function is exposed to the client-side JavaScript via google.script.run.
 * @param {string} userMessage The message received from the user via the chat interface.
 * @returns {Object} Un obiect { message: string, buttons: Array<string>|null }.
 */
function processUserMessage(userMessage) {
  Logger.log(`Received user message from frontend: ${userMessage}`);

  let finalResponse = "A apărut o problemă necunoscută. Te rog încearcă din nou.";
  let responseButtons = null;

  const userSheets = getOrCreateAndSetUserSheetIds();
  if (!userSheets) {
    Logger.log("Eroare: ID-urile foilor utilizatorului nu au putut fi setate. Acces refuzat sau eroare internă.");
    return { message: "Îmi pare rău, a apărut o problemă la inițializarea datelor utilizatorului. Te rog încearcă mai târziu.", buttons: null };
  }

  // >>> Modificare cheie aici: Verificăm starea conversației înainte de a apela AI-ul
  const conversationState = getConversationState();
  Logger.log(`DEBUG: Starea conversației actuală: ${JSON.stringify(conversationState)}`);

  if (conversationState.state === "awaiting_client_cui_and_email") {
    Logger.log("DEBUG: Stare 'awaiting_client_cui_and_email' detectată. Apelăm handleInvoiceIssuanceFlow.");
    const response = handleInvoiceIssuanceFlow(userMessage, conversationState.data);
    finalResponse = response.message;
    responseButtons = response.buttons;
    // Starea este ștearsă sau re-setată în handleInvoiceIssuanceFlow
  } else {
    // Dacă nu există o stare anterioară sau starea nu e relevantă, apelăm ChatGPT pentru o nouă intenție
    const aiRawResponse = callChatGPT(userMessage);

    try {
      const parsedResponse = JSON.parse(aiRawResponse);
      Logger.log(`DEBUG: Parsed AI response JSON: ${JSON.stringify(parsedResponse)}`);

      if (parsedResponse.intent === "check_stock") {
        Logger.log(`DEBUG: Detected 'check_stock' intent.`);
        const idProdus = parsedResponse.product_id;
        if (!idProdus) {
          finalResponse = "Ce produs vrei să verifici stocul?";
        } else {
          const product = getProductDataByID_Produs(idProdus);
          if (product) {
            finalResponse = `Stocul disponibil ${product.nume} (ID: ${product.idProdus}) este: ${product.inventar}.`;
          } else {
            finalResponse = `Îmi pare rău, nu am putut identifica produsul cu ID: ${idProdus}. Te rog verifică ID-ul și încearcă din nou.`;
          }
        }
        clearConversationState(); // Terminăm fluxul de verificare stoc
      } else if (parsedResponse.intent === "add_order") {
        Logger.log("DEBUG: Detected 'add_order' intent.");
        const products = parsedResponse.products;
        const clientCui = parsedResponse.customer_id;
        const clientEmail = parsedResponse.customer_email;

        // Validare preliminară: verifică stocul pentru TOATE produsele înainte de a iniția fluxul
        let allProductsInStock = true;
        if (products && products.length > 0) {
          for (const item of products) {
            const product = getProductDataByID_Produs(item.product_id);
            if (!product) {
              finalResponse = `Produsul cu ID-ul '${item.product_id}' nu a fost găsit. Te rog verifică și încearcă din nou.`;
              allProductsInStock = false;
              break;
            }
            // Asigură-te că item.quantity este o valoare numerică validă
            const requestedQuantity = parseInt(item.quantity);
            if (isNaN(requestedQuantity) || requestedQuantity <= 0) {
              finalResponse = `Cantitatea solicitată pentru produsul '${product.nume}' este invalidă. Te rog specifică o cantitate numerică pozitivă.`;
              allProductsInStock = false;
              break;
            }
            if (parseInt(product.inventar) < requestedQuantity) {
              finalResponse = `Stoc insuficient pentru ${product.nume}. Stoc actual: ${product.inventar}. Cantitate solicitată: ${requestedQuantity}.`;
              allProductsInStock = false;
              break;
            }
          }
        } else {
          // Niciun produs valid detectat în cererea de comandă
          finalResponse = "Nu am putut identifica produsele și cantitățile pentru comandă. Te rog specifică clar. Ex: 'comanda 10 tricouri'.";
          allProductsInStock = false;
        }

        if (allProductsInStock) {
          if (clientCui && clientEmail) {
            // Toate detaliile sunt complete și stocul este OK, procesează comanda direct
            finalResponse = handleOrderCreation(products, clientCui, clientEmail);
            clearConversationState(); // Comanda finalizată, ștergem starea
          } else {
            // Lipsesc CUI/Email, setăm starea și cerem detalii
            setConversationState("awaiting_client_cui_and_email", { products: products }); // Salvăm detaliile produselor
            finalResponse = "Pentru a finaliza comanda, am nevoie de CUI-ul și adresa de email ale clientului. Le poți introduce acum? (ex: '47315510 test@example.com')";
          }
        }
      } else if (parsedResponse.intent === "missing_order_details" && parsedResponse.message) {
        finalResponse = parsedResponse.message;
      } else if (parsedResponse.intent === "unknown_product" && parsedResponse.query) {
        finalResponse = `Îmi pare rău, nu am putut găsi produsul "${parsedResponse.query}". Te rog verifică numele și încearcă din nou.`;
        clearConversationState();
      } else if (parsedResponse.intent === "get_total_products") {
        const totalProducts = getTotalProducts();
        finalResponse = (totalProducts !== null) ? `Numărul total de produse înregistrate este: ${totalProducts}.` : "Îmi pare rău, nu am putut recupera numărul total de produse.";
        clearConversationState();
      } else if (parsedResponse.intent === "get_total_stock_value") {
        const totalStockValue = getTotalStockValue();
        finalResponse = (totalStockValue !== null) ? `Valoarea totală a stocului este: ${totalStockValue.toFixed(2)} RON.` : "Îmi pare rău, nu am putut recupera valoarea totală a stocului.";
        clearConversationState();
      } else if (parsedResponse.intent === "get_num_orders") {
        const numOrders = getNumberOfOrders();
        finalResponse = (numOrders !== null) ? `Numărul total de comenzi înregistrate este: ${numOrders}.` : "Îmi pare rău, nu am putut recupera numărul total de comenzi.";
        clearConversationState();
      } else if (parsedResponse.intent === "get_total_order_revenue") {
        const totalRevenue = getTotalOrderRevenue();
        finalResponse = (totalRevenue !== null) ? `Venitul total din comenzi este: ${totalRevenue.toFixed(2)} RON.` : "Îmi pare rău, nu am putut recupera venitul total din comenzi.";
        clearConversationState();
      } else if (parsedResponse.intent === "get_low_stock_count") {
        const lowStockProducts = getLowStockCount();
        if (lowStockProducts !== null) {
          finalResponse = (lowStockProducts.length > 0) ? `Produsele cu stoc sub ${LOW_STOCK_THRESHOLD} sunt:\n- ${lowStockProducts.join('\n- ')}.` : `Nu există produse cu stoc sub ${LOW_STOCK_THRESHOLD}. Stocul este optim!`;
        } else {
          finalResponse = "Îmi pare rău, nu am putut recupera lista de produse cu stoc scăzut.";
        }
        clearConversationState();
      } else if (parsedResponse.intent === "get_capabilities" && parsedResponse.capabilities) {
        Logger.log(`DEBUG: Detected 'get_capabilities' intent.`);
        const capabilitiesList = parsedResponse.capabilities;
        if (capabilitiesList.length > 0) {
          finalResponse = `Pot face următoarele:\n- ${capabilitiesList.join('\n- ')}.`;
        } else {
          finalResponse = "Îmi pare rău, nu am putut identifica capabilitățile mele.";
        }
        clearConversationState();
      } else { // Fallback pentru orice altă intenție nerecunoscută
        finalResponse = "Îmi pare rău, nu am înțeles cererea. Te rog să încerci din nou sau să folosești un format precum: 'comanda 10 tricouri pentru 47315510 test@email.com'.";
        Logger.log(`DEBUG: Intenție nerecunoscută sau altă cerere: ${userMessage}`);
        clearConversationState(); // Nu există o stare complexă de menținut
      }
    } catch (e) {
      Logger.log(`Eroare la procesarea mesajului (apel ChatGPT): ${e.message}. Răspuns AI brut: ${aiRawResponse}`);
      finalResponse = `A apărut o eroare la procesarea solicitării. Te rog încearcă din nou.`;
      clearConversationState();
    }
  }

  Logger.log(`Trimitere răspuns final către frontend: ${finalResponse}, Butoane: ${JSON.stringify(responseButtons)}`);
  return { message: finalResponse, buttons: responseButtons };
}




//==============================================================================
// DATA ACCESS FUNCTIONS (CRUD Operations on User Sheets)
//==============================================================================

/**
 * Retrieves product details by ID_Produs from the Produse sheet.
 * @param {string} idProdus The ID_Produs to search for.
 * @returns {Object|null} An object containing product name, description, inventory, prices (excl/incl TVA), and Cota_TVA, or null if not found.
 */
function getProductDataByID_Produs(idProdus) {
  try {
    const ss = SpreadsheetApp.openById(CURRENT_USER_PRODUSE_ID);
    const sheet = ss.getSheetByName("Produse Data");
    if (!sheet) {
      Logger.log("Sheet 'Produse Data' not found in Produse spreadsheet.");
      return null;
    }

    const data = sheet.getDataRange().getValues();
    // New Headers: ID_Produs(0), Nume_Produs(1), Descriere_Produs(2), Cantitate_Produs(3), Pret_unitar_achizitie_fara_TVA(4), Pret_unitar_vanzare_fara_TVA(5), Cota_TVA(6), Pret_unitar_vanzare_cu_TVA(7)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idProdus) {
        return {
          idProdus: data[i][0],
          nume: data[i][1],
          descriere: data[i][2],
          inventar: data[i][3],
          pretAchizitieFaraTVA: data[i][4],
          pretVanzareFaraTVA: data[i][5],
          cotaTVA: data[i][6],
          pretVanzareCuTVA: data[i][7]
        };
      }
    }
    Logger.log(`Product with ID_Produs '${idProdus}' not found.`);
    return null;
  } catch (e) {
    Logger.log(`Error in getProductDataByID_Produs: ${e.message}`);
    return null;
  }
}


/**
 * Writes a new order entry to the Comenzi Data sheet.
 * @param {string} idComanda Unique ID for the order.
 * @param {string} orderDate Date of the order (YYYY-MM-DD).
 * @param {string} clientCui Client's CUI (replaces idClient).
 * @param {string} idProdus Product ID.
 * @param {number} quantity Quantity of the product ordered.
 * @param {number} unitPriceNoVat Unit price of the product without VAT (Pret_unitar_fara_TVA_RON).
 * @param {number} itemValueNoVat The value of the item without VAT (quantity * unitPriceNoVat).
 * @param {string} paymentTermDate The payment term date (YYYY-MM-DD).
 * @param {string} statusFactura Initial status (e.g., "In asteptare").
 * @returns {boolean} True if successful, false otherwise.
 */
function createNewOrderEntry(idComanda, orderDate, clientCui, idProdus, quantity, unitPriceNoVat, itemValueNoVat, paymentTermDate, statusFactura) {
  try {
    const ss = SpreadsheetApp.openById(CURRENT_USER_COMENZI_ID);
    const sheet = ss.getSheetByName("Comenzi Data");
    if (!sheet) {
      Logger.log("Sheet 'Comenzi Data' not found in Comenzi spreadsheet.");
      return false;
    }

    const vatRate = DEFAULT_VAT_RATE;

    const valoareRon = itemValueNoVat;
    const valoareTVA_RON = valoareRon * vatRate;
    const totalPlataRon = valoareRon + valoareTVA_RON;

    // NEW Comenzi Headers: ID_Comanda(0), ID_Client(1), ID_Produs(2), Cantitate_Produs(3), Pret_unitar_fara_TVA_RON(4),
    // Valoare_TVA_RON(5), Total_Plata_RON(6), Data_Comanda(7), Termen_de_Plata(8), Status_Factura(9)
    const newRow = [
      idComanda,          // Index 0
      clientCui,          // Index 1
      idProdus,           // Index 2
      quantity,           // Index 3
      unitPriceNoVat,     // Index 4
      valoareTVA_RON,     // Index 5
      totalPlataRon,      // Index 6
      orderDate,          // Index 7
      paymentTermDate,    // Index 8
      statusFactura       // Index 9
    ];

    sheet.appendRow(newRow);
    Logger.log(`New order ${idComanda} added successfully.`);
    return true;
  } catch (e) {
    Logger.log(`Error creating new order entry: ${e.message}`);
    return false;
  }
}

/**
 * Updates the stock level for a product in the Produse Data sheet.
 * @param {string} idProdus The ID_Produs of the product to update.
 * @param {number} quantityChange The amount to change stock by (e.g., -5 for decrease, +10 for increase).
 * @returns {boolean} True if successful, false otherwise.
 */
function updateProductStock(idProdus, quantityChange) {
  try {
    const ss = SpreadsheetApp.openById(CURRENT_USER_PRODUSE_ID);
    const sheet = ss.getSheetByName("Produse Data");
    if (!sheet) {
      Logger.log("Sheet 'Produse Data' not found in Produse spreadsheet.");
      return false;
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const stockColumnIndex = headers.indexOf("Cantitate_Produs");

    if (stockColumnIndex === -1) {
      Logger.log("Error: 'Cantitate_Produs' column not found in Produse Data sheet.");
      return false;
    }

    let foundRow = -1;
    let currentStock = 0;

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === idProdus) {
        foundRow = i;
        currentStock = parseInt(data[i][stockColumnIndex]);
        break;
      }
    }

    if (foundRow === -1) {
      Logger.log(`Product with ID_Produs '${idProdus}' not found for stock update.`);
      return false;
    }

    const newStock = currentStock + quantityChange;

    if (newStock < 0) {
      Logger.log(`Error: Insufficient stock for product ${idProdus}. Current: ${currentStock}, Attempted change: ${quantityChange}.`);
      return false;
    }

    sheet.getRange(foundRow + 1, stockColumnIndex + 1).setValue(newStock);
    Logger.log(`Stock for ${idProdus} updated from ${currentStock} to ${newStock}.`);
    return true;
  } catch (e) {
    Logger.log(`Error updating stock for product ${idProdus}: ${e.message}`);
    return false;
  }
}

//==============================================================================
// CORE BUSINESS LOGIC (Order Creation, Invoice Generation)
//==============================================================================

/**
 * Handles the creation of a new order, including validation, stock update, and invoice generation for multiple products.
 * @param {Array<Object>} products A list of product objects, each with {product_id: string, quantity: number}.
 * @param {string} clientCui The client's CUI (fiscal code).
 * @param {string} clientEmail The client's email address.
 * @returns {string} A user-friendly message indicating success or failure.
 */
function handleOrderCreation(products, clientCui, clientEmail) {
  Logger.log(`Intrat în handleOrderCreation pentru ${products.length} produse, CUI: ${clientCui}, Email: ${clientEmail}`);

  // 1. Obține detalii client de la ANAF (o singură dată per comandă)
  const clientDetailsFromANAF = getCompanyInfoFromANAF(clientCui);
  if (!clientDetailsFromANAF) {
    return `Îmi pare rău, nu am putut găsi detalii pentru clientul cu CUI "${clientCui}" în baza de date ANAF. Te rog verifică CUI-ul.`;
  }
  const customer = {
      nume: clientDetailsFromANAF.denumire,
      cif: clientDetailsFromANAF.cui,
      regCom: clientDetailsFromANAF.nrRegCom,
      adresa: clientDetailsFromANAF.adresa,
      judet: "", // ANAF API doesn't directly provide Judet in date_generale, needs parsing from adresa or separate lookup
      tara: "Romania",
      email: clientEmail, // Use email provided by user
      telefon: ""
  };

  // 2. Generare ID Comandă și Date (o singură dată per comandă)
  const orderId = "COM" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss") + Math.floor(Math.random() * 1000);
  const orderDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const paymentTermDate = Utilities.formatDate(new Date(new Date().setDate(new Date().getDate() + 30)), Session.getScriptTimeZone(), "yyyy-MM-dd");

  let totalOrderValueNoVat = 0; // Va acumula valoarea totală a comenzii (fără TVA)
  let successfullyProcessedProducts = []; // Pentru a urmări ce produse au fost procesate cu succes

  // 3. Iterare prin produse, creare intrare și actualizare stoc pentru FIECARE produs
  for (const item of products) {
    const productId = item.product_id;
    const quantity = item.quantity;

    const product = getProductDataByID_Produs(productId);
    if (!product) {
      Logger.log(`Produsul cu ID-ul '${productId}' nu a fost găsit. Sărim peste acest produs.`);
      continue;
    }

    const unitPrice = parseFloat(product.pretVanzareFaraTVA);
    const itemValueNoVat = quantity * unitPrice;

    const orderSuccess = createNewOrderEntry(
      orderId,
      orderDate,
      clientCui,
      productId,
      quantity,
      unitPrice,
      itemValueNoVat,
      paymentTermDate,
      "In asteptare"
    );

    if (orderSuccess) {
      const stockUpdateSuccess = updateProductStock(productId, -quantity);
      if (stockUpdateSuccess) {
        totalOrderValueNoVat += itemValueNoVat;
        successfullyProcessedProducts.push({
            product: product,
            quantity: quantity,
            itemValueNoVat: itemValueNoVat
        });
        Logger.log(`Produs ${productId} (Cant: ${quantity}) procesat cu succes. Stoc nou: ${parseInt(product.inventar) - quantity}.`);
      } else {
        Logger.log(`Comanda înregistrată pentru ${productId}, dar eroare la actualizarea stocului.`);
      }
    } else {
      Logger.log(`Eroare la înregistrarea comenzii pentru produsul ${productId}.`);
    }
  }

  if (successfullyProcessedProducts.length === 0) {
    return `Îmi pare rău, nu am putut procesa niciun produs din comandă. Te rog verifică detaliile.`;
  }

  // 4. Generare și trimitere factură (o singură dată pentru întreaga comandă)
  const invoiceResult = generateAndEmailInvoice(
    orderId,
    orderDate,
    successfullyProcessedProducts,
    customer,
    totalOrderValueNoVat,
    paymentTermDate
  );

  const orderFinalTotals = calculateOrderTotals(totalOrderValueNoVat, DEFAULT_VAT_RATE);
  const finalTotalPaymentWithVat = orderFinalTotals.totalPaymentWithVat;

      let finalOrderMessage = `Comanda #${orderId} a fost înregistrată cu succes pentru ${customer.nume} (${clientCui}).`;
      finalOrderMessage += ` Total de plata: ${finalTotalPaymentWithVat.toFixed(2)} RON.`;

      if (invoiceResult.success) {
          finalOrderMessage += ` Factura a fost trimisă pe email la ${customer.email}.`;
      } else {
          finalOrderMessage += ` A apărut o problemă la generarea sau trimiterea facturii pe email.`;
      }

      finalOrderMessage += ` Stocul a fost actualizat.`;

      return finalOrderMessage;
    }

//==============================================================================
// AUTOMATED PROCESSES
//==============================================================================

/**
       * Generates an invoice PDF from a template and emails it to the customer, without storing copies in Drive.
       * @param {string} orderId The ID of the order.
       * @param {string} orderDate The date of the order.
       * @param {Array<Object>} productList A list of product objects, each with {product: Object, quantity: number, itemValueNoVat: number}.
       * @param {Object} customerDetails Object with customer: {name, email, adresa, cif, regCom, judet, tara}.
       * @param {number} totalOrderValueNoVat Total value of ALL items in the order without VAT.
       * @param {string} paymentTermDate The date payment is due.
       * @returns {Object} {success: boolean, totalPaymentWithVat: number|null}
       */
    function generateAndEmailInvoice(orderId, orderDate, productList, customerDetails, totalOrderValueNoVat, paymentTermDate) {
      let tempInvoiceFile = null;
      try {
        Logger.log("DEBUG: Începe generarea facturii.");
        const templateFile = DriveApp.getFileById(INVOICE_TEMPLATE_ID);
        Logger.log(`DEBUG: Șablon factură obținut cu ID: ${INVOICE_TEMPLATE_ID}`);

        const newInvoiceName = `Factura_${orderId}_TEMP.docx`;
        tempInvoiceFile = templateFile.makeCopy(newInvoiceName, DriveApp.getRootFolder());
        Logger.log(`DEBUG: Copie temporară creată: ${tempInvoiceFile.getName()} cu ID: ${tempInvoiceFile.getId()}`);

        const doc = DocumentApp.openById(tempInvoiceFile.getId());
        const body = doc.getBody();
        Logger.log("DEBUG: Documentul temporar deschis.");

        const companyDetails = getCompanyDetails();
        if (!companyDetails) {
            Logger.log("Eroare: Detaliile companiei nu au fost găsite în foaia Date firma. Nu se poate genera factura.");
            return { success: false, totalPaymentWithVat: null };
        }
        Logger.log(`DEBUG: Detalii companie obținute: ${companyDetails.nume}`);

        let totalVatValue = 0;
        let totalPaymentWithVat = 0;

        let itemNrContent = [];
        let productNameContent = [];
        let umContent = [];
        let quantityContent = [];
        let unitPriceContent = [];
        let valueContent = [];
        let vatValueContent = [];

        let itemNrCounter = 1;
        for (const item of productList) {
            const product = item.product;
            const quantity = item.quantity;
            const itemValueNoVat = item.itemValueNoVat;

            let vatRate = parseFloat(product.cotaTVA);
            if (isNaN(vatRate)) {
                Logger.log(`Avertisment: Cota_TVA (${product.cotaTVA}) invalidă pentru produsul ${product.idProdus}. Se utilizează rata TVA implicită.`);
                vatRate = DEFAULT_VAT_RATE;
            }

            const itemVatValue = itemValueNoVat * vatRate;
            const itemUnitPriceDisplay = parseFloat(product.pretVanzareFaraTVA).toFixed(2);

            totalVatValue += itemVatValue;
            totalPaymentWithVat += (itemValueNoVat + itemVatValue);

            itemNrContent.push(itemNrCounter.toString());
            productNameContent.push(product.nume);
            umContent.push("buc");
            quantityContent.push(quantity.toString());
            unitPriceContent.push(itemUnitPriceDisplay);
            valueContent.push(itemValueNoVat.toFixed(2));
            vatValueContent.push(itemVatValue.toFixed(2));

            itemNrCounter++;
        }

        const finalItemNr = itemNrContent.join('\n');
        const finalProductName = productNameContent.join('\n');
        const finalUm = umContent.join('\n');
        const finalQuantity = quantityContent.join('\n');
        const finalUnitPrice = unitPriceContent.join('\n');
        const finalValue = valueContent.join('\n');
        const finalVatValue = vatValueContent.join('\n');

        const tables = body.getTables();
        let productTable = null;
        for (let i = 0; i < tables.length; i++) {
            const table = tables[i];
            if (table.getNumRows() > 0 && table.getCell(0, 0).getText().trim() === "Nr. crt") {
                productTable = table;
                break;
            }
        }

        if (!productTable) {
            Logger.log("Error: Product table with 'Nr. crt' header not found in the invoice template. Cannot populate items.");
            return { success: false, totalPaymentWithVat: null };
        }

        const targetRow = productTable.getRow(2);
        if (targetRow.getNumCells() !== 7) {
            Logger.log(`Eroare: Rândul țintă al tabelului de produse (index 2) are ${targetRow.getNumCells()} celule, se așteptau 7. Nu se pot popula elementele corect.`);
            return { success: false, totalPaymentWithVat: null };
        }

        targetRow.getCell(0).setText(finalItemNr);
        targetRow.getCell(1).setText(finalProductName);
        targetRow.getCell(2).setText(finalUm);
        targetRow.getCell(3).setText(finalQuantity);
        targetRow.getCell(4).setText(finalUnitPrice);
        targetRow.getCell(5).setText(finalValue);
        targetRow.getCell(6).setText(finalVatValue);
        Logger.log(`DEBUG: Celule populate cu produse multiple în rânduri noi.`);

        const formattedOrderDate = Utilities.formatDate(new Date(orderDate), Session.getScriptTimeZone(), "dd/MM/yyyy");
        const formattedPaymentTermDate = Utilities.formatDate(new Date(paymentTermDate), Session.getScriptTimeZone(), "dd/MM/yyyy");
        const vatRateDisplay = (productList.length > 0 && productList[0].product && !isNaN(parseFloat(productList[0].product.cotaTVA))) ? (parseFloat(productList[0].product.cotaTVA) * 100).toFixed(0) : (DEFAULT_VAT_RATE * 100).toFixed(0);


        body.replaceText('{{INVOICE_SERIES}}', companyDetails.nume.substring(0, 2).toUpperCase());
        body.replaceText('{{INVOICE_NUMBER}}', orderId);
        body.replaceText('{{INVOICE_DATE}}', formattedOrderDate);
        body.replaceText('{{PAYMENT_TERM_DATE}}', formattedPaymentTermDate);
        body.replaceText('{{VAT_RATE_DISPLAY}}', vatRateDisplay);

        body.replaceText('{{COMPANY_NAME}}', companyDetails.nume);
        body.replaceText('{{COMPANY_CIF}}', companyDetails.cif);
        body.replaceText('{{COMPANY_REG_COM}}', companyDetails.regCom);
        body.replaceText('{{COMPANY_ADDRESS}}', companyDetails.adresa);
        body.replaceText('{{COMPANY_COUNTY}}', companyDetails.judet);
        body.replaceText('{{COMPANY_IBAN_RON}}', companyDetails.ibanRon);
        body.replaceText('{{COMPANY_BANK}}', companyDetails.banca);
        body.replaceText('{{COMPANY_CAPITAL}}', companyDetails.capital);
        body.replaceText('{{COMPANY_PHONE}}', companyDetails.telefon);
        body.replaceText('{{COMPANY_EMAIL}}', companyDetails.email);
        body.replaceText('{{PREPARER_NAME}}', companyDetails.nume);

        body.replaceText('{{CUSTOMER_NAME}}', customerDetails.nume);
        body.replaceText('{{CUSTOMER_CIF}}', customerDetails.cif);
        body.replaceText('{{CUSTOMER_REG_COM}}', customerDetails.regCom);
        body.replaceText('{{CUSTOMER_ADDRESS}}', customerDetails.adresa);
        body.replaceText('{{CUSTOMER_COUNTY}}', customerDetails.judet);
        body.replaceText('{{CUSTOMER_COUNTRY}}', customerDetails.tara);
        body.replaceText('{{CUSTOMER_EMAIL}}', customerDetails.email);

        body.replaceText('{{TOTAL_VALUE_NO_VAT}}', totalOrderValueNoVat.toFixed(2));
        body.replaceText('{{TOTAL_VAT_VALUE}}', totalVatValue.toFixed(2));
        body.replaceText('{{TOTAL_PAYMENT_WITH_VAT}}', totalPaymentWithVat.toFixed(2));

        doc.saveAndClose();

        const pdfBlob = tempInvoiceFile.getAs(MimeType.PDF);
        pdfBlob.setName(`Factura_${orderId}.pdf`);
        Logger.log("DEBUG: PDF Blob creat.");

        GmailApp.sendEmail(
          customerDetails.email,
          `Factura #${orderId} pentru comanda dumneavoastră Orderly`,
          `Bună ziua, ${customerDetails.nume},\n\nVă atașăm factura pentru comanda dumneavoastră #${orderId}.\n\nVă mulțumim!\nEchipa Orderly`,
          {
            attachments: [pdfBlob],
            name: "Orderly AI Assistant"
          }
        );
        Logger.log(`DEBUG: Email trimis către ${customerDetails.email}.`);

        if (companyDetails.emailContabil && companyDetails.emailContabil.trim() !== '') {
            GmailApp.sendEmail(
                companyDetails.emailContabil,
                `Copie Factura #${orderId} - ${customerDetails.nume}`,
                `Bună ziua,\n\nVă atașăm copia facturii #${orderId} pentru clientul ${customerDetails.nume}.\n\nCu respect,\nOrderly AI Assistant`,
                {
                    attachments: [pdfBlob],
                    name: "Orderly AI Assistant"
                }
            );
            Logger.log(`DEBUG: Copie factură trimisă contabilului: ${companyDetails.emailContabil}.`);
        }

              Logger.log(`Invoice ${orderId} generată și trimisă prin email către ${customerDetails.email}.`);
          return { success: true, totalPaymentWithVat: totalPaymentWithVat };
        } catch (e) {
          Logger.log(`Eroare la generarea sau trimiterea facturii pentru comanda ${orderId}: ${e.message}`);
          return { success: false, totalPaymentWithVat: null };
        } finally {
        if (tempInvoiceFile) {
          try {
            tempInvoiceFile.setTrashed(true);
            Logger.log(`Fișierul temporar de factură ${tempInvoiceFile.getName()} mutat la coș.`);
          } catch (e) {
            Logger.log(`Eroare la mutarea fișierului temporar la coș: ${e.message}`);
          }
        }
      }
    }


/**
 * Reads company details from the user's dedicated "Date firma" sheet.
 * @returns {Object|null} Object containing company details, or null if sheet/data not found.
 */
function getCompanyDetails() {
  try {
    const ss = SpreadsheetApp.openById(CURRENT_USER_DATE_FIRMA_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("No company data found in Date firma sheet.");
      return null;
    }
    const headers = data[0];
    const companyRow = data[1];

    return {
      nume: companyRow[headers.indexOf("Nume_Firma")],
      cif: companyRow[headers.indexOf("CIF_Firma")],
      regCom: companyRow[headers.indexOf("Reg_Com_Firma")],
      adresa: companyRow[headers.indexOf("Adresa_Firma")],
      judet: companyRow[headers.indexOf("Judet_Firma")],
      ibanRon: companyRow[headers.indexOf("IBAN_Firma")],
      banca: companyRow[headers.indexOf("Banca_Firma")],
      capital: companyRow[headers.indexOf("Capital_Firma")],
      telefon: companyRow[headers.indexOf("Telefon_Firma")],
      email: companyRow[headers.indexOf("Email_firma")],
      emailContabil: companyRow[headers.indexOf("Email_contabil")]
    };
  } catch (e) {
    Logger.log(`Error in getCompanyDetails: ${e.message}`);
    return null;
  }
}


//==============================================================================
// KPI CALCULATION FUNCTIONS
//==============================================================================

/**
 * Calculates the total number of products directly from the Produse Data sheet.
 * @returns {number|null} The total product count, or null if an error occurs.
 */
function getTotalProducts() {
  try {
    const ss = SpreadsheetApp.openById(CURRENT_USER_PRODUSE_ID);
    const sheet = ss.getSheetByName("Produse Data");
    if (!sheet) {
      Logger.log("Produse Data sheet not found.");
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const totalProducts = data.length > 0 ? data.length - 1 : 0;

    Logger.log(`Total Produse calculated directly: ${totalProducts}`);
    return totalProducts;
  } catch (e) {
    Logger.log(`Error in getTotalProducts (direct calculation): ${e.message}`);
    return null;
  }
}

/**
 * Calculates the total monetary value of all products in stock.
 * @returns {number|null} The total stock value, or null if an error occurs.
 */
function getTotalStockValue() {
  try {
    const ss = SpreadsheetApp.openById(CURRENT_USER_PRODUSE_ID);
    const sheet = ss.getSheetByName("Produse Data");
    if (!sheet) {
      Logger.log("Produse Data sheet not found for stock value calculation.");
      return null;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("No product data found for stock value calculation.");
      return 0;
    }

    let totalValue = 0;
    for (let i = 1; i < data.length; i++) {
      const stock = parseFloat(data[i][3]);
      const price = parseFloat(data[i][5]);

      if (!isNaN(stock) && !isNaN(price)) {
        totalValue += stock * price;
      } else {
        Logger.log(`Warning: Non-numeric data found in row ${i + 1} of Produse Data for stock or selling price.`);
      }
    }

    Logger.log(`Total Stock Value calculated directly: ${totalValue}`);
    return totalValue;
  } catch (e) {
    Logger.log(`Error in getTotalStockValue (direct calculation): ${e.message}`);
    return null;
  }
}

/**
 * Calculates the total number of orders directly from the Comenzi Data sheet.
 * @returns {number|null} The total order count, or null if an error occurs.
 */
function getNumberOfOrders() {
  try {
    const ss = SpreadsheetApp.openById(CURRENT_USER_COMENZI_ID);
    const sheet = ss.getSheetByName("Comenzi Data");
    if (!sheet) {
      Logger.log("Comenzi Data sheet not found for order count calculation.");
      return null;
    }

    const data = sheet.getDataRange().getValues();
    const numOrders = data.length > 0 ? data.length - 1 : 0;

    Logger.log(`Number of Orders calculated directly: ${numOrders}`);
    return numOrders;
  } catch (e) {
    Logger.log(`Error in getNumberOfOrders (direct calculation): ${e.message}`);
    return null;
  }
}

/**
 * Calculates the total monetary revenue from all orders.
 * @returns {number|null} The total order revenue, or null if an error occurs.
 */
function getTotalOrderRevenue() {
  try {
    const ss = SpreadsheetApp.openById(CURRENT_USER_COMENZI_ID);
    const sheet = ss.getSheetByName("Comenzi Data");
    if (!sheet) {
      Logger.log("Comenzi Data sheet not found for revenue calculation.");
      return null;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("No order data found for revenue calculation.");
      return 0;
    }

    let totalRevenue = 0;
    for (let i = 1; i < data.length; i++) {
      const orderTotal = parseFloat(data[i][6]);
      if (!isNaN(orderTotal)) {
        totalRevenue += orderTotal;
      } else {
        Logger.log(`Warning: Non-numeric data found in row ${i + 1}, column G (Total_Plata_RON) of Comenzi Data for total revenue.`);
      }
    }

    Logger.log(`Total Order Revenue calculated directly: ${totalRevenue}`);
    return totalRevenue;
  } catch (e) {
    Logger.log(`Error in getTotalOrderRevenue (direct calculation): ${e.message}`);
    return null;
  }
}


/**
 * Finds and returns the names and stock levels of products with stock below a defined threshold.
 * @returns {Array<string>|null} An array of strings like "Product Name (Stock: X)", or null if an error occurs.
 */
function getLowStockCount() {
  try {
    const ss = SpreadsheetApp.openById(CURRENT_USER_PRODUSE_ID);
    const sheet = ss.getSheetByName("Produse Data");
    if (!sheet) {
      Logger.log("Produse Data sheet not found for low stock calculation.");
      return null;
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("No product data found for low stock calculation.");
      return [];
    }

    let lowStockProducts = [];
    for (let i = 1; i < data.length; i++) {
      const productName = data[i][1];
      const stock = parseFloat(data[i][3]);

      if (!isNaN(stock) && stock < LOW_STOCK_THRESHOLD) {
        lowStockProducts.push(`${productName} (Stoc: ${stock})`);
      }
    }

    Logger.log(`Low Stock Products found: ${JSON.stringify(lowStockProducts)}`);
    return lowStockProducts;
  } catch (e) {
    Logger.log(`Error in getLowStockCount (getting names): ${e.message}`);
    return null;
  }
}




//==============================================================================
// EXTERNAL API INTEGRATIONS
//==============================================================================
/**
 * Interrogates the ANAF API to retrieve company information based on CUI.
 * @param {string} cui The fiscal code (CUI) of the company to search for.
 * @returns {Object|null} An object containing requested company details (cui, data, denumire, adresa, nrRegCom) or null if not found/error.
 */
function getCompanyInfoFromANAF(cui) {
  const ANAF_API_URL = "https://webservicesp.anaf.ro/api/PlatitorTvaRest/v9/tva";
  const queryDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const numericCui = cui.startsWith('RO') ? cui.substring(2) : cui;

  const requestBody = JSON.stringify([
    {
      "cui": numericCui,
      "data": queryDate
    }
  ]);

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": requestBody,
    "muteHttpExceptions": true
  };

  Logger.log(`Sending request to ANAF API for CUI: ${cui} on date: ${queryDate}`);

  try {
    const response = UrlFetchApp.fetch(ANAF_API_URL, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    Logger.log(`ANAF API Response Code: ${responseCode}`);
    Logger.log(`ANAF API Response Body: ${responseBody}`);

    if (responseCode === 200) {
      const jsonResponse = JSON.parse(responseBody);

      if (jsonResponse.found && jsonResponse.found.length > 0) {
        const companyData = jsonResponse.found[0].date_generale;
        Logger.log(`ANAF found company: ${companyData.denumire}`);

        return {
          cui: companyData.cui,
          data: companyData.data,
          denumire: companyData.denumire,
          adresa: companyData.adresa,
          nrRegCom: companyData.nrRegCom
        };
      } else {
        Logger.log(`ANAF API: CUI ${cui} not found in response, or unexpected structure.`);
        return null;
      }
    } else {
      Logger.log(`ANAF API Error: Code ${responseCode}, Body: ${responseBody}`);
      return null;
    }
  } catch (e) {
    Logger.log(`Exception calling ANAF API for CUI ${cui}: ${e.message}`);
    return null;
  }
}




//==============================================================================
// HELPER FUNCTIONS
//==============================================================================

/**
 * Calculează valorile totale ale comenzii (TVA și Total de plată).
 * @param {number} totalValueNoVat Valoarea totală a comenzii fără TVA.
 * @param {number} vatRate Rata TVA aplicabilă (ex: 0.19 sau 0.21).
 * @returns {Object} Un obiect cu { totalVatValue: number, totalPaymentWithVat: number }.
 */
function calculateOrderTotals(totalValueNoVat, vatRate) {
  const totalVatValue = totalValueNoVat * vatRate;
  const totalPaymentWithVat = totalValueNoVat + totalVatValue;
  return {
    totalVatValue: totalVatValue,
    totalPaymentWithVat: totalPaymentWithVat
  };
}




function forceUrlFetchAuth() {
  try {
    UrlFetchApp.fetch("https://www.google.com"); // Un apel simplu, nevinovat, care necesită permisiunea
    Logger.log("UrlFetchApp authorization triggered successfully.");
  } catch (e) {
    Logger.log("Error forcing UrlFetchApp auth: " + e.message);
  }
}
