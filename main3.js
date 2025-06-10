
// --- START OF FILE main3.js ---


// --- START OF FILE main3.js ---

// --- START OF FILE main3.js ---

// --- START OF FILE main3.js ---

// --- START OF FULL main3.js ---

const LOCAL_STORAGE_INDEX_KEY = 'fullFileManagerIndexData';
const DRIVE_SYNC_INTERVAL = 30000; // 30 seconds

// ADDED: IndexedDB Helper to handle large data caching and avoid localStorage quota errors.
const dbHelper = {
    db: null,
    dbName: 'FileContentDB',
    storeName: 'fileChunks',
    init: function() {
        return new Promise((resolve, reject) => {
            if (this.db) {
                return resolve(this.db);
            }
            if (!window.indexedDB) {
                return reject("IndexedDB not supported by this browser.");
            }
            const request = indexedDB.open(this.dbName, 1);

            request.onerror = (event) => {
                console.error("IndexedDB error:", request.error);
                reject("IndexedDB error: " + request.error);
            };

            request.onsuccess = (event) => {
                this.db = event.target.result;
                console.log("IndexedDB initialized successfully.");
                resolve(this.db);
            };

            request.onupgradeneeded = (event) => {
                console.log("Upgrading IndexedDB...");
                const db = event.target.result;
                if (!db.objectStoreNames.contains(this.storeName)) {
                    db.createObjectStore(this.storeName);
                    console.log(`Object store "${this.storeName}" created.`);
                }
            };
        });
    },
    set: function(key, value) {
        return new Promise((resolve, reject) => {
            if (!this.db) return reject("DB not initialized");
            const transaction = this.db.transaction([this.storeName], 'readwrite');
            const store = transaction.objectStore(this.storeName);
            const request = store.put(value, key);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });
    },
    get: function(key) {
        return new Promise((resolve, reject) => {
            if (!this.db) return reject("DB not initialized");
            const transaction = this.db.transaction([this.storeName], 'readonly');
            const store = transaction.objectStore(this.storeName);
            const request = store.get(key);
            request.onsuccess = () => resolve(request.result); // Will be undefined if not found
            request.onerror = () => reject(request.error);
        });
    },
    delete: function(key) {
        return new Promise((resolve, reject) => {
            if (!this.db) return reject("DB not initialized");
            const transaction = this.db.transaction([this.storeName], 'readwrite');
            const store = transaction.objectStore(this.storeName);
            const request = store.delete(key);
            request.onsuccess = () => resolve();
            request.onerror = () => reject(request.error);
        });
    }
};

// MSAL Configuration
const msalConfig = {
    auth: {
        clientId: MSAL_CLIENT_ID, // From env.js
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin,
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
    }
};
const msalInstance = new msal.PublicClientApplication(msalConfig);
const msalScopes = ["Files.ReadWrite", "User.Read"];
let microsoftAccount = null;
let oneDriveFragmentFolderId = null; // This stores the ID of the ONEDRIVE_FRAGMENT_FOLDER_NAME

// DOM Elements
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const uploadProgress = document.getElementById('uploadProgress');
const breadcrumbContainer = document.getElementById('breadcrumbContainer');
const tabFileTreeBtn = document.getElementById('tabFileTree');
const tabDataTableBtn = document.getElementById('tabDataTable');
const contentFileTree = document.getElementById('contentFileTree');
const contentDataTable = document.getElementById('contentDataTable');
const fileTreeContainer = document.getElementById('fileTreeContainer');
const newFolderBtn = document.getElementById('newFolderBtn');
const searchInput = document.getElementById('searchInput');
const dataTableBody = document.getElementById('dataTableBody');
let filesDataTableHeaders;
const itemsPerPageFilesSelect = document.getElementById('itemsPerPageFilesSelect');
const paginationControlsFiles = document.getElementById('paginationControlsFiles');
const prevPageBtnFiles = document.getElementById('prevPageBtnFiles');
const nextPageBtnFiles = document.getElementById('nextPageBtnFiles');
const pageInfoFiles = document.getElementById('pageInfoFiles');
const metadataModal = document.getElementById('metadataModal');
const modalFileId = document.getElementById('modalFileId');
const modalFileName = document.getElementById('modalFileName');
const modalFileTags = document.getElementById('modalFileTags');
const modalFileComments = document.getElementById('modalFileComments');
const modalFilePublicAccess = document.getElementById('modalFilePublicAccess');
const modalPublicLinkStatus = document.getElementById('modalPublicLinkStatus');
const modalSaveBtn = document.getElementById('modalSaveBtn');
const modalCancelBtn = document.getElementById('modalCancelBtn');
const modalDeleteBtn = document.getElementById('modalDeleteBtn');
const newFolderModal = document.getElementById('newFolderModal');
const modalNewFolderName = document.getElementById('modalNewFolderName');
const modalParentFolder = document.getElementById('modalParentFolder');
const modalNewFolderCreateBtn = document.getElementById('modalNewFolderCreateBtn');
const modalNewFolderCancelBtn = document.getElementById('modalNewFolderCancelBtn');
const renameFolderModal = document.getElementById('renameFolderModal');
const modalRenameFolderId = document.getElementById('modalRenameFolderId');
const modalRenameFolderNameInput = document.getElementById('modalRenameFolderNameInput');
const modalRenameFolderParentInfo = document.getElementById('modalRenameFolderParentInfo');
const modalRenameFolderSaveBtn = document.getElementById('modalRenameFolderSaveBtn');
const modalRenameFolderCancelBtn = document.getElementById('modalRenameFolderCancelBtn');
const modalDeleteFolderBtn = document.getElementById('modalDeleteFolderBtn');
const tabSharersBtn = document.getElementById('tabSharers');
const contentSharers = document.getElementById('contentSharers');
const newSharerBtn = document.getElementById('newSharerBtn');
const sharersSearchInput = document.getElementById('sharersSearchInput');
const sharersTableBody = document.getElementById('sharersTableBody');
let sharersDataTableHeaders;
const itemsPerPageSharersSelect = document.getElementById('itemsPerPageSharersSelect');
const paginationControlsSharers = document.getElementById('paginationControlsSharers');
const prevPageBtnSharers = document.getElementById('prevPageBtnSharers');
const nextPageBtnSharers = document.getElementById('nextPageBtnSharers');
const pageInfoSharers = document.getElementById('pageInfoSharers');
const sharerModal = document.getElementById('sharerModal');
const sharerModalTitle = document.getElementById('sharerModalTitle');
const modalSharerId = document.getElementById('modalSharerId');
const modalSharerShortname = document.getElementById('modalSharerShortname');
const modalSharerFullname = document.getElementById('modalSharerFullname');
const modalSharerEmail = document.getElementById('modalSharerEmail');
const modalSharerDept = document.getElementById('modalSharerDept');
const modalSharerType = document.getElementById('modalSharerType');
const modalSharerPassword = document.getElementById('modalSharerPassword');
const modalSharerStopAllShares = document.getElementById('modalSharerStopAllShares');
const modalSharerSaveBtn = document.getElementById('modalSharerSaveBtn');
const modalSharerCancelBtn = document.getElementById('modalSharerCancelBtn');
const modalSharerDeleteBtn = document.getElementById('modalSharerDeleteBtn');
const modalFileShareSearch = document.getElementById('modalFileShareSearch');
const modalFileShareSearchResults = document.getElementById('modalFileShareSearchResults');
const modalFileCurrentlySharedWith = document.getElementById('modalFileCurrentlySharedWith');
const modalFileNoSharersMsg = document.getElementById('modalFileNoSharersMsg');
const modalFolderShareSearch = document.getElementById('modalFolderShareSearch');
const modalFolderShareSearchResults = document.getElementById('modalFolderShareSearchResults');
const modalFolderCurrentlySharedWith = document.getElementById('modalFolderCurrentlySharedWith');
const modalFolderNoSharersMsg = document.getElementById('modalFolderNoSharersMsg');
const toggleUploadAreaBtn = document.getElementById('toggleUploadAreaBtn');
const uploadAreaWrapper = document.getElementById('uploadAreaWrapper');
const viewFileModal = document.getElementById('viewFileModal');
const viewFileModalContent = document.getElementById('viewFileModalContent');
const viewFileModalTitle = document.getElementById('viewFileModalTitle');
const viewFileModalCloseBtn = document.getElementById('viewFileModalCloseBtn');
const viewFileIframe = document.getElementById('viewFileIframe');
const googleApiStatusEl = document.getElementById('googleApiStatus');
const googleSignInBtn = document.getElementById('googleSignInBtn');
const microsoftSignInBtn = document.getElementById('microsoftSignInBtn');
const microsoftApiStatusEl = document.getElementById('microsoftApiStatus');

// App State
let files = [];
let folders = [];
let expandedFolders = ['root'];
let activeFolderId = 'root';
let filesDataTablePage = 1;
let filesSortColumn = 'currentName';
let filesSortDirection = 'asc';
let itemsPerPageFiles = 10;
let sharers = [];
let sharersTablePage = 1;
let sharersSortColumn = 'fullname';
let sharersSortDirection = 'asc';
let itemsPerPageSharers = 10;
let shares = [];
let currentFileModalShares = [];
let currentFolderModalShares = [];
let draggedItemId = null;
let draggedItemType = null;
let deepLinkProcessed = false; // Flag to ensure deep link is processed once

// Google Drive State
let tokenClient, googleAccessToken = null, gDriveUploadFolderId = null; // Renamed accessToken to googleAccessToken
let gDriveIndexFileId = null;
let isDriveAuthenticated = false;
let isDriveReadyForOps = false;
let isIndexDataDirty = false;
let backgroundSyncIntervalId = null;
let lastSuccessfulSyncTimestamp = 0;

// Microsoft State
let isMicrosoftAuthenticated = false;

// --- Microsoft Authentication and OneDrive Functions ---
async function msLogin() {
    try {
        await msalInstance.initialize();
        const accounts = msalInstance.getAllAccounts();
        if (accounts.length > 0) {
            msalInstance.setActiveAccount(accounts[0]);
            microsoftAccount = accounts[0];
        } else {
            const loginResponse = await msalInstance.loginPopup({ scopes: msalScopes });
            msalInstance.setActiveAccount(loginResponse.account);
            microsoftAccount = loginResponse.account;
        }

        if (microsoftAccount) {
            isMicrosoftAuthenticated = true;
            if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `✅ Signed in to Microsoft as ${ microsoftAccount.username }.`;
            if (microsoftSignInBtn) {
                microsoftSignInBtn.textContent = `✅ MS: ${ microsoftAccount.name || microsoftAccount.username.split('@')[0] } `;
                microsoftSignInBtn.disabled = true;
            }
            const token = await getMicrosoftAccessToken();
            if (token) {
                await ensureOneDriveFolderExists(token, ONEDRIVE_FRAGMENT_FOLDER_NAME);
                if (oneDriveFragmentFolderId) {
                    if (microsoftApiStatusEl) microsoftApiStatusEl.textContent += ` Using OneDrive folder: "${ONEDRIVE_FRAGMENT_FOLDER_NAME}".`;
                } else {
                    if (microsoftApiStatusEl) microsoftApiStatusEl.textContent += ` Error ensuring OneDrive folder "${ONEDRIVE_FRAGMENT_FOLDER_NAME}".Check console.`;
                }
            } else {
                if (microsoftApiStatusEl) microsoftApiStatusEl.textContent += ` Could not get MS token to verify folder.`;
            }
        }
    } catch (err) {
        console.error("MSAL Login/Init Error: ", err);
        if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `❌ Microsoft Login failed: ${ err.message || err } `;
        isMicrosoftAuthenticated = false;
    } finally {
        await checkAndProcessDeepLink(); // Process deep link after MS auth attempt
    }
}

async function getMicrosoftAccessToken() {
    if (!microsoftAccount) {
        microsoftAccount = msalInstance.getActiveAccount();
        if (!microsoftAccount && msalInstance.getAllAccounts().length > 0) {
            microsoftAccount = msalInstance.getAllAccounts()[0];
            msalInstance.setActiveAccount(microsoftAccount);
        }
    }
    if (!microsoftAccount) {
        console.warn("Microsoft user not signed in. Cannot acquire token.");
        if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = "⚠️ MS Sign-in required for token.";
        return null;
    }
    const request = { scopes: msalScopes, account: microsoftAccount };
    try {
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
    } catch (error) {
        console.warn("Silent token acquisition failed for MS Graph, attempting popup:", error);
        if (error instanceof msal.InteractionRequiredAuthError || error instanceof msal.BrowserAuthError) {
            try {
                const response = await msalInstance.acquireTokenPopup(request);
                return response.accessToken;
            } catch (popupError) {
                console.error("Popup token acquisition failed for MS Graph:", popupError);
                if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `❌ MS Token Error: ${ popupError.message || popupError } `;
                return null;
            }
        } else {
            console.error("Other MS Graph token acquisition error:", error);
            if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `❌ MS Token Error: ${ error.message || error } `;
            return null;
        }
    }
}

async function ensureOneDriveFolderExists(token, folderName) {
    if (!token) {
        console.error("ensureOneDriveFolderExists: No MS token provided.");
        if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = "❌ OneDrive Error: Missing token for folder check.";
        return null;
    }
    const folderPathUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(folderName)}`;
try {
    const res = await fetch(folderPathUrl, {
        method: "GET",
        headers: { Authorization: `Bearer ${token}` }
    });

    if (res.ok) {
        const folderData = await res.json();
        if (folderData && folderData.folder) {
            console.log(`OneDrive folder "${folderName}" found with ID: ${folderData.id}`);
            oneDriveFragmentFolderId = folderData.id; // Correctly assigns to the global variable
            return folderData.id;
        } else {
            console.warn(`Item "${folderName}" exists at path but is not a folder. Will attempt to create.`);
        }
    } else if (res.status !== 404) {
        const errData = await res.json().catch(() => ({ error: { message: "Unknown error parsing failed response." } }));
        throw new Error(`Failed to check/get folder by path: ${res.status} ${errData.error?.message || res.statusText}`);
    }

    console.log(`OneDrive folder "${folderName}" not found or item at path is not a folder, attempting to create it...`);
    const createUrl = `https://graph.microsoft.com/v1.0/me/drive/root/children`;
    const createRes = await fetch(createUrl, {
        method: "POST",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
        body: JSON.stringify({
            name: folderName,
            folder: {},
            "@microsoft.graph.conflictBehavior": "fail"
        })
    });

    if (!createRes.ok) {
        const errData = await createRes.json().catch(() => ({ error: { message: "Unknown error parsing failed create response." } }));
        throw new Error(`Failed to create folder: ${createRes.status} ${errData.error?.message || createRes.statusText}`);
    }
    const newFolder = await createRes.json();
    console.log(`OneDrive folder "${folderName}" created with ID: ${newFolder.id}`);
    oneDriveFragmentFolderId = newFolder.id; // Correctly assigns to the global variable
    return newFolder.id;

} catch (error) {
    console.error("Error ensuring OneDrive folder exists:", error);
    if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `❌ OneDrive folder error: ${error.message}`;
    oneDriveFragmentFolderId = null;
    return null;
}
}

async function uploadFragmentToOneDrive(localFileId, fragmentDataString) {
    if (!isMicrosoftAuthenticated) {
        console.warn("Microsoft not authenticated. Cannot upload fragment to OneDrive.");
        throw new Error("Microsoft authentication required for fragment upload.");
    }
    if (!oneDriveFragmentFolderId) {
        const tempTokenForFolderCheck = await getMicrosoftAccessToken();
        if (tempTokenForFolderCheck) {
            await ensureOneDriveFolderExists(tempTokenForFolderCheck, ONEDRIVE_FRAGMENT_FOLDER_NAME);
        }
        if (!oneDriveFragmentFolderId) {
            if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `❌ OneDrive Error: Fragment folder ID for "${ONEDRIVE_FRAGMENT_FOLDER_NAME}" could not be established.`;
            throw new Error(`Could not obtain OneDrive fragment folder ID for "${ONEDRIVE_FRAGMENT_FOLDER_NAME}".`);
        }
    }

    const token = await getMicrosoftAccessToken();
    if (!token) {
        throw new Error("Could not get Microsoft access token for fragment upload.");
    }

    const fragmentJsonContent = {
        fileid: localFileId,
        fragment: String(fragmentDataString.length),
        fragmentPosition: "last",
        fragmentData: fragmentDataString
    };
    const blob = new Blob([JSON.stringify(fragmentJsonContent, null, 2)], { type: "application/json" });
    const oneDriveFileName = `${localFileId}_frag.json`;
    const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${oneDriveFragmentFolderId}:/${encodeURIComponent(oneDriveFileName)}:/content`;

    if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `Uploading fragment ${oneDriveFileName} to OneDrive...`;
    console.log(`Attempting to upload fragment to OneDrive: ${uploadUrl}`);

    try {
        const response = await fetch(uploadUrl, {
            method: "PUT",
            headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
            body: blob
        });

        if (response.ok) {
            const result = await response.json();
            if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `✅ Fragment ${result.name} uploaded to OneDrive.`;
            console.log(`Fragment ${result.name} uploaded to OneDrive. ID: ${result.id}`);
            return result.id;
        } else {
            const errData = await response.json().catch(() => ({ error: { message: response.statusText } }));
            const errText = errData.error?.message || response.statusText;
            throw new Error(`OneDrive Upload Error (${response.status}): ${errText}`);
        }
    } catch (err) {
        console.error("OneDrive Fragment Upload Error (catch block):", err);
        if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `❌ OneDrive Fragment Upload failed: ${err.message}`;
        throw err;
    }
}

// --- ADDED FUNCTION to upload password to OneDrive ---
async function uploadPasswordToOneDrive(localFileId, passwordString) {
    if (!isMicrosoftAuthenticated) {
        console.warn("Microsoft not authenticated. Cannot upload password to OneDrive.");
        throw new Error("Microsoft authentication required for password upload.");
    }
    if (!oneDriveFragmentFolderId) { // Use the same folder as fragments for simplicity
        const tempTokenForFolderCheck = await getMicrosoftAccessToken();
        if (tempTokenForFolderCheck) {
            await ensureOneDriveFolderExists(tempTokenForFolderCheck, ONEDRIVE_FRAGMENT_FOLDER_NAME);
        }
        if (!oneDriveFragmentFolderId) {
            if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `❌ OneDrive Error: Folder ID for "${ONEDRIVE_FRAGMENT_FOLDER_NAME}" (for password) could not be established.`;
            throw new Error(`Could not obtain OneDrive folder ID for password storage in "${ONEDRIVE_FRAGMENT_FOLDER_NAME}".`);
        }
    }

    const token = await getMicrosoftAccessToken();
    if (!token) {
        throw new Error("Could not get Microsoft access token for password upload.");
    }

    const passwordJsonContent = {
        fileid: localFileId,
        password: passwordString,
        version: "1.0.0" // For future compatibility
    };
    const blob = new Blob([JSON.stringify(passwordJsonContent, null, 2)], { type: "application/json" });
    const oneDriveFileName = `${localFileId}_pwd.json`; // Differentiate from fragment
    const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${oneDriveFragmentFolderId}:/${encodeURIComponent(oneDriveFileName)}:/content`;

    if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `Uploading password file ${oneDriveFileName} to OneDrive...`;
    console.log(`Attempting to upload password file to OneDrive: ${uploadUrl}`);

    try {
        const response = await fetch(uploadUrl, {
            method: "PUT",
            headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
            body: blob
        });

        if (response.ok) {
            const result = await response.json();
            if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `✅ Password file ${result.name} uploaded to OneDrive.`;
            console.log(`Password file ${result.name} uploaded to OneDrive. ID: ${result.id}`);
            return result.id;
        } else {
            const errData = await response.json().catch(() => ({ error: { message: response.statusText } }));
            const errText = errData.error?.message || response.statusText;
            throw new Error(`OneDrive Password Upload Error (${response.status}): ${errText}`);
        }
    } catch (err) {
        console.error("OneDrive Password Upload Error (catch block):", err);
        if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `❌ OneDrive Password Upload failed: ${err.message}`;
        throw err; // Re-throw to be caught by calling function
    }
}


async function deleteOneDriveFileByItemId(itemId) { // Removed token from args, will get it inside
    if (!isMicrosoftAuthenticated) {
        console.warn("Microsoft not authenticated. Cannot delete OneDrive item.");
        return false;
    }
    const token = await getMicrosoftAccessToken();
    if (!itemId || !token) {
        console.warn("deleteOneDriveFileByItemId: Missing itemId or token.");
        return false;
    }
    const url = `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}`;
    console.log(`Attempting to delete OneDrive item: ${url}`);
    try {
        const response = await fetch(url, {
            method: 'DELETE',
            headers: { Authorization: `Bearer ${token}` }
        });
        if (response.ok || response.status === 204) {
            console.log(`File item ${itemId} deleted from OneDrive.`);
            if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `Item ${itemId} deleted from OneDrive.`;
            return true;
        } else {
            const errData = await response.json().catch(() => ({ error: { message: `Status ${response.statusText}` } }));
            const errMsg = errData.error?.message || response.statusText;
            console.error(`Failed to delete OneDrive item ${itemId}: ${response.status}`, errMsg);
            if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `❌ Error deleting item ${itemId} from OneDrive: ${errMsg}`;
            return false;
        }
    } catch (error) {
        console.error(`Exception deleting OneDrive item ${itemId}:`, error);
        if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = `❌ Exception deleting item ${itemId} from OneDrive.`;
        return false;
    }
}

// --- Google Drive Core Functions ---
function authenticateGoogle(forcePrompt = false) {
    const tokenOptions = {
        prompt: forcePrompt ? 'consent' : '',
        hint: localStorage.getItem('googleUserEmail') || ''
    };
    tokenClient.requestAccessToken(tokenOptions);
}

// Check browser
// When you receive the token:
if (typeof GOOGLE_CLIENT_ID === 'undefined') {
    console.error('GOOGLE_CLIENT_ID is not defined. Check that env.js is loaded correctly.');
}

const isUserSafari = isSafari();
function isSafari() {
    return /^((?!chrome|android).)*safari/i.test(navigator.userAgent);
}


try {
    if (isUserSafari) {
        // Use redirect-based flow for Safari
        window.location.href = `https://accounts.google.com/o/oauth2/v2/auth?client_id=${GOOGLE_CLIENT_ID}&redirect_uri=${REDIRECT_URI}&response_type=token&scope=email%20profile%20openid&prompt=consent`;
    } else {
        // Use initTokenClient for all other browsers
        const tokenClient = google.accounts.oauth2.initTokenClient({
            client_id: GOOGLE_CLIENT_ID,
            scope: 'email profile openid',
            callback: (tokenResponse) => {
                // handle token response
                console.log('Token response:', tokenResponse);
            }
        });

        // You still need this to be called from a click event
        // tokenClient.requestAccessToken(); // This should be called on user interaction
    }
} catch (error) {
    console.error('OAuth flow failed:', error);
    // Optionally show user-friendly message
    // alert('Something went wrong while signing in. Please try again.'); // Commented out as it might be too intrusive at this stage
}


try {
    if (window.google?.accounts?.oauth2) {
        tokenClient = google.accounts.oauth2.initTokenClient({
            client_id: GOOGLE_CLIENT_ID,
            scope: 'https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.metadata.readonly',
            callback: async (tokenResponse) => {
                if (tokenResponse && tokenResponse.access_token) {
                    googleAccessToken = tokenResponse.access_token;
                    localStorage.setItem('googleAccessToken', tokenResponse.access_token);
                    localStorage.setItem('googleAccessTokenTimestamp', Date.now().toString());
                    isDriveAuthenticated = true;
                    console.log("Google authentication successful via token client callback.");
                    gapi.load('client', initializeGapiClient);
                } else {
                    console.warn("Google token response missing access_token. Attempting fallback.");
                    localStorage.setItem('googleAccessTokenTimestamp', Date.now().toString());
                    await checkAndProcessDeepLink();
                }
            },
        });
    } else {
        console.warn("Google OAuth2 API is not available. Skipping token client init.");
    }
} catch (error) {
    console.error('Token client setup failed:', error);
}



async function initializeGapiClient() {
    if (googleApiStatusEl) googleApiStatusEl.textContent = "Initializing Google Drive client...";
    try {
        await gapi.client.init({});
    } catch (err) {
        console.error("GAPI client init error:", err);
        if (googleApiStatusEl) googleApiStatusEl.textContent = "GAPI client init error.";
        isDriveReadyForOps = false;
        await checkAndProcessDeepLink(); // Attempt deep link even if GAPI init fails
        return;
    }

    try {
        await gapi.client.load('drive', 'v3');
        console.log("Google Drive API v3 loaded.");
    } catch (e) {
        console.error("Error loading Google Drive API v3:", e);
        if (googleApiStatusEl) googleApiStatusEl.textContent = "Error loading Drive API. Permissions may not work.";
        isDriveReadyForOps = false;
        await checkAndProcessDeepLink(); // Attempt deep link even if API load fails (for local items)
        return;
    }
    gapi.client.setToken({ access_token: googleAccessToken });
    // isDriveAuthenticated should already be true if we got here through token callback
    if (googleApiStatusEl) googleApiStatusEl.textContent = "Google Drive client initialized. Setting up application folder...";
    if (googleSignInBtn) {
        googleSignInBtn.textContent = "🔄 Syncing...";
        googleSignInBtn.disabled = true;
    }
    await ensureUploadFolderAndManageIndex(); // This will eventually call checkAndProcessDeepLink
}

async function _internalFetchAndApplyIndexFromDrive() {
    if (!gDriveIndexFileId) {
        console.warn("_internalFetchAndApplyIndexFromDrive: No GDrive Index File ID, cannot fetch.");
        if (googleApiStatusEl) googleApiStatusEl.textContent = "Error: Drive Index File ID missing. Cannot confirm sync.";
        return false;
    }
    console.log("Attempting to re-fetch index from Drive...");
    try {
        const indexContentRes = await gapi.client.request({
            path: `/drive/v3/files/${gDriveIndexFileId}?alt=media`, method: 'GET'
        });

        let indexDataText;
        if (typeof indexContentRes.body === 'string') {
            indexDataText = indexContentRes.body;
        } else if (typeof indexContentRes.body === 'object' && indexContentRes.body !== null) {
            indexDataText = JSON.stringify(indexContentRes.body);
        } else {
            throw new Error("Unexpected index file content type from Drive on re-fetch.");
        }

        const parsedDriveData = JSON.parse(indexDataText);
        files = parsedDriveData.files || [];
        folders = parsedDriveData.folders || [];
        expandedFolders = parsedDriveData.expandedFolders || ['root'];
        activeFolderId = parsedDriveData.activeFolderId || 'root';
        sharers = parsedDriveData.sharers || [];
        shares = parsedDriveData.shares || [];

        if (!folders.find(f => f.id === 'root')) {
            folders.unshift({ id: 'root', name: 'Root', parentId: null, isDeleted: false, deletedDate: null });
            if (!expandedFolders.includes('root')) expandedFolders.push('root');
        }
        files.forEach(f => {
            f.isDeleted = !!f.isDeleted;
            f.driveFileId = f.driveFileId || null;
            f.isPublic = !!f.isPublic;
            f.encryptedDriveJsonId = f.encryptedDriveJsonId || null;
            f.oneDriveFragmentId = f.oneDriveFragmentId || null;
            f.oneDrivePasswordItemId = f.oneDrivePasswordItemId || null; // Handle new field
        });
        folders.forEach(f => { f.isDeleted = !!f.isDeleted; });

        const currentActiveFolderIsValid = folders.find(f => f.id === activeFolderId && !f.isDeleted);
        if (!currentActiveFolderIsValid) activeFolderId = 'root';

        localStorage.setItem(LOCAL_STORAGE_INDEX_KEY, JSON.stringify(parsedDriveData));
        if (googleApiStatusEl) googleApiStatusEl.textContent = `Index re-fetched from Drive. Last sync: ${new Date().toLocaleTimeString()}`;
        console.log("Applied Drive index to local state:", parsedDriveData);
        lastSuccessfulSyncTimestamp = Date.now();
        return true;
    } catch (error) {
        console.error("Error re-fetching or applying index from Drive:", error);
        if (googleApiStatusEl) googleApiStatusEl.textContent = `Error confirming sync: ${error.message || 'Fetch failed'}. Local data might be ahead of Drive.`;
        return false;
    }
}

async function ensureUploadFolderAndManageIndex() { // Google Drive folder
    try {
        // DEFAULT_GDRIVE_UPLOAD_FOLDER_NAME from env.js
        const qParams = `name='${GDRIVE_UPLOAD_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
        const folderSearchRes = await gapi.client.drive.files.list({
            q: qParams,
            fields: 'files(id, name)',
            spaces: 'drive'
        });

        const driveFolders = folderSearchRes.result.files;
        if (driveFolders.length > 0) {
            gDriveUploadFolderId = driveFolders[0].id;
        } else {
            const fileMetadata = {
                name: GDRIVE_UPLOAD_FOLDER_NAME,
                mimeType: 'application/vnd.google-apps.folder'
            };
            const createFolderRes = await gapi.client.drive.files.create({ resource: fileMetadata, fields: 'id' });
            gDriveUploadFolderId = createFolderRes.result.id;
            if (googleApiStatusEl) googleApiStatusEl.textContent = `App folder "${GDRIVE_UPLOAD_FOLDER_NAME}" created on Drive.`;
        }

        if (!gDriveUploadFolderId) throw new Error("Could not obtain Google Drive upload folder ID.");
        if (googleApiStatusEl) googleApiStatusEl.textContent = `Using Drive folder: ${GDRIVE_UPLOAD_FOLDER_NAME}. Checking for index file...`;

        const indexSearchRes = await gapi.client.drive.files.list({
            q: `'${gDriveUploadFolderId}' in parents and name='dms_index.json' and trashed=false`,
            fields: 'files(id, name)',
            spaces: 'drive'
        });
        const indexFiles = indexSearchRes.result.files;

        if (indexFiles.length > 0) {
            gDriveIndexFileId = indexFiles[0].id;
            if (googleApiStatusEl) googleApiStatusEl.textContent = "Found index file on Drive. Downloading and applying...";
            const fetchSuccess = await _internalFetchAndApplyIndexFromDrive();
            if (!fetchSuccess) throw new Error("Failed to fetch and apply initial index from Drive.");

            if (sharers.length === 0 && !localStorage.getItem('hasInitializedExampleSharers')) {
                initializeExampleSharers();
                localStorage.setItem('hasInitializedExampleSharers', 'true');
                const dataWithExampleSharers = { files, folders, expandedFolders, activeFolderId, sharers, shares };
                await _doUploadIndexDataToDrive(dataWithExampleSharers);
                const reFetchSuccess = await _internalFetchAndApplyIndexFromDrive();
                if (!reFetchSuccess) console.warn("Failed to re-fetch after adding example sharers.");
            }
            if (googleApiStatusEl && lastSuccessfulSyncTimestamp > 0) googleApiStatusEl.textContent = `Data successfully synced from Google Drive. Last sync: ${new Date(lastSuccessfulSyncTimestamp).toLocaleTimeString()}`;
            else if (googleApiStatusEl) googleApiStatusEl.textContent = `Data successfully synced from Google Drive.`;

        } else {
            if (googleApiStatusEl) googleApiStatusEl.textContent = "No index file on Drive. Using local data (if any) or defaults, then uploading new index.";
            loadDataFromLocalStorage(); // Load local first (redundant if called before, but safe)
            const localDataToUpload = { files, folders, expandedFolders, activeFolderId, sharers, shares };
            await _doUploadIndexDataToDrive(localDataToUpload); // Upload it

            const fetchAfterCreateSuccess = await _internalFetchAndApplyIndexFromDrive(); // Fetch it back
            if (!fetchAfterCreateSuccess) {
                throw new Error("Uploaded new index, but failed to re-fetch and apply it.");
            }
            if (googleApiStatusEl && lastSuccessfulSyncTimestamp > 0) googleApiStatusEl.textContent = `Initial index created on Drive and synced. Last sync: ${new Date(lastSuccessfulSyncTimestamp).toLocaleTimeString()}`;
            else if (googleApiStatusEl) googleApiStatusEl.textContent = `Initial index created on Drive and synced.`;
        }
        isDriveReadyForOps = true;
        isIndexDataDirty = false; // Reset dirty flag after successful sync/setup
        startBackgroundSync();
        renderAll(); // Render after data is confirmed
        if (googleSignInBtn) googleSignInBtn.textContent = "✅ Synced with Google";

    } catch (error) {
        console.error("Error during Drive folder/index setup:", error);
        if (googleApiStatusEl) googleApiStatusEl.textContent = `Drive Error: ${error.message || 'Setup failed'}. Using local data. Some Drive features may be disabled.`;
        isDriveReadyForOps = false;
        loadDataFromLocalStorage(); // Fallback to local data
        renderAll(); // Render with local data
        if (googleSignInBtn) googleSignInBtn.textContent = "⚠️ Drive Sync Issue";
    } finally {
        if (googleSignInBtn) googleSignInBtn.disabled = false;
        await checkAndProcessDeepLink(); // Crucial: process deep link after GDrive setup attempt (success or fail)
    }
}

async function _doUploadIndexDataToDrive(indexData) { // Google Drive
    if (!isDriveAuthenticated || !gDriveUploadFolderId) {
        if (googleApiStatusEl) googleApiStatusEl.textContent = "Cannot upload index: Not authenticated or Drive folder not ready.";
        console.warn("_doUploadIndexDataToDrive: Pre-conditions not met.");
        throw new Error("Drive not ready for index upload.");
    }
    if (googleApiStatusEl) googleApiStatusEl.textContent = "Syncing index to Google Drive...";
    console.log("Uploading index data to Drive:", indexData);
    const blob = new Blob([JSON.stringify(indexData, null, 2)], { type: 'application/json' });
    const metadata = { name: 'dms_index.json', mimeType: 'application/json' };
    if (!gDriveIndexFileId) { // Creating new index file
        metadata.parents = [gDriveUploadFolderId];
    }

    const form = new FormData();
    form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
    form.append('file', blob);

    const method = gDriveIndexFileId ? 'PATCH' : 'POST';
    const path = gDriveIndexFileId ? `/upload/drive/v3/files/${gDriveIndexFileId}` : `/upload/drive/v3/files`;
    const params = { uploadType: 'multipart' };
    if (!gDriveIndexFileId) {
        params.fields = 'id'; // Only need ID if creating
    }
    const url = `https://www.googleapis.com${path}?${new URLSearchParams(params).toString()}`;

    try {
        const res = await fetch(url, {
            method, headers: { Authorization: 'Bearer ' + googleAccessToken }, body: form
        });
        const result = await res.json();
        if (res.ok) {
            if (!gDriveIndexFileId && result.id) { // If created, store new ID
                gDriveIndexFileId = result.id;
                console.log("Index created on Drive, new ID:", gDriveIndexFileId);
            }
            if (googleApiStatusEl) googleApiStatusEl.textContent = "Index successfully uploaded to Google Drive.";
            console.log("Index uploaded/updated on Drive:", result);
        } else {
            throw new Error(result.error?.message || "Unknown error uploading index");
        }
    } catch (err) {
        console.error("Error uploading index file:", err);
        if (googleApiStatusEl) googleApiStatusEl.textContent = `Error syncing index: ${err.message}. Data saved locally.`;
        throw err;
    }
}

async function uploadActualFileToDrive(fileObject, driveFileName, originalFileNameForProgress) { // Google Drive
    if (!isDriveReadyForOps) throw new Error("Drive not ready for file operations.");
    if (uploadProgress) uploadProgress.textContent = `Starting upload of ${originalFileNameForProgress} to Drive...`;
    const CHUNK_SIZE = 10 * 1024 * 1024; // 10MB
    const isLarge = fileObject.size > CHUNK_SIZE;
    const driveFileMetadata = { name: driveFileName, parents: [gDriveUploadFolderId] };

    if (!isLarge) {
        const form = new FormData();
        form.append('metadata', new Blob([JSON.stringify(driveFileMetadata)], { type: 'application/json' }));
        form.append('file', fileObject);
        const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType,size', {
            method: 'POST', headers: { Authorization: 'Bearer ' + googleAccessToken }, body: form
        });
        if (!res.ok) { const errData = await res.json(); throw new Error(`Simple upload failed: ${errData.error?.message || res.statusText}`); }
        if (uploadProgress) uploadProgress.textContent = `${originalFileNameForProgress} uploaded to Drive.`;
        return await res.json();
    } else {
        const resumableInitMetadata = { name: driveFileName, parents: [gDriveUploadFolderId] };
        const sessionRes = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&fields=id,name,mimeType,size', {
            method: 'POST',
            headers: {
                Authorization: 'Bearer ' + googleAccessToken,
                'Content-Type': 'application/json; charset=UTF-8',
                'X-Upload-Content-Type': fileObject.type || 'application/octet-stream',
                'X-Upload-Content-Length': fileObject.size.toString()
            },
            body: JSON.stringify(resumableInitMetadata)
        });

        if (!sessionRes.ok) { const errData = await sessionRes.json(); throw new Error(`Resumable session start failed: ${errData.error?.message || sessionRes.statusText}`); }
        const sessionUrl = sessionRes.headers.get('Location');
        if (!sessionUrl) throw new Error("Resumable session URL not received.");

        let offset = 0;
        let finalDriveFileObject = null;
        while (offset < fileObject.size) {
            const chunk = fileObject.slice(offset, Math.min(offset + CHUNK_SIZE, fileObject.size));
            const uploadChunkRes = await fetch(sessionUrl, { // Session URL is pre-authorized
                method: 'PUT',
                headers: { 'Content-Range': `bytes ${offset}-${offset + chunk.size - 1}/${fileObject.size}` },
                body: chunk
            });
            if (uploadProgress) uploadProgress.textContent = `Uploading ${originalFileNameForProgress} to Drive: ${Math.round(((offset + chunk.size) / fileObject.size) * 100)}%`;

            if (uploadChunkRes.status === 200 || uploadChunkRes.status === 201) { // Done
                finalDriveFileObject = await uploadChunkRes.json();
                break;
            } else if (uploadChunkRes.status === 308) { // In progress
                const rangeHeader = uploadChunkRes.headers.get('Range');
                if (rangeHeader) {
                    const match = rangeHeader.match(/bytes=0-(\d+)/);
                    if (match && parseInt(match[1], 10) + 1 > offset) {
                        offset = parseInt(match[1], 10) + 1;
                    } else { // If no Range or unexpected, just advance by chunk size
                        offset += chunk.size;
                    }
                } else { // If no Range header, advance by chunk size (less precise but a fallback)
                    offset += chunk.size;
                }
            } else { // Error
                const errData = await uploadChunkRes.json().catch(() => ({}));
                throw new Error(`Resumable chunk upload failed (${uploadChunkRes.status}): ${errData.error?.message || uploadChunkRes.statusText}`);
            }
        }
        if (!finalDriveFileObject) throw new Error("Resumable upload finished but no file object retrieved.");
        if (uploadProgress) uploadProgress.textContent = `${originalFileNameForProgress} uploaded to Drive.`;
        return finalDriveFileObject;
    }
}

async function deleteDriveFileById(driveFileId) { // Google Drive
    if (!isDriveReadyForOps || !driveFileId) {
        console.warn("Cannot delete from Drive: Not ready or no Drive File ID.", { isDriveReadyForOps, driveFileId });
        return; // Return undefined or throw error, depending on desired strictness
    }
    try {
        await gapi.client.drive.files.delete({ fileId: driveFileId });
        console.log(`File ${driveFileId} deleted from Google Drive.`);
        if (googleApiStatusEl) googleApiStatusEl.textContent = `A file was deleted from Drive.`;
    } catch (error) {
        console.error(`Failed to delete file ${driveFileId} from Google Drive:`, error.result?.error?.message || error);
        if (googleApiStatusEl) googleApiStatusEl.textContent = `Error deleting a file from Drive. It might need manual deletion.`;
        throw error; // Re-throw to allow caller to know it failed
    }
}

// --- Core Local App Functions ---
function generateUUID() { if (typeof crypto !== 'undefined' && crypto.randomUUID) { return crypto.randomUUID(); } else { return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) { var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8); return v.toString(16); }); } }
function _commitLocalStateChanges() { const dataToSave = { files, folders, expandedFolders, activeFolderId, sharers, shares }; localStorage.setItem(LOCAL_STORAGE_INDEX_KEY, JSON.stringify(dataToSave)); isIndexDataDirty = true; console.log("_commitLocalStateChanges: Local data saved, index marked dirty."); }
function initializeDefaultStateAndSharers() { initializeDefaultState(); if (!localStorage.getItem('hasInitializedExampleSharers')) { initializeExampleSharers(); localStorage.setItem('hasInitializedExampleSharers', 'true'); isIndexDataDirty = true; } }
function initializeExampleSharers() { sharers = [{ id: 'sharer_ex_001', shortname: 'alice_g', fullname: 'Alice Green', email: 'alice@example.com', dept: 'Marketing', type: 'internal', password: 'password123', stopAllShares: false, isExample: true }, { id: 'sharer_ex_002', shortname: 'bob_r', fullname: 'Bob Red', email: 'bob.red@external.com', dept: 'Client XYZ', type: 'external', password: 'password123', stopAllShares: false, isExample: true }, { id: 'sharer_ex_003', shortname: 'charlie_b', fullname: 'Charlie Blue', email: 'charlie@example.com', dept: 'Engineering', type: 'internal', password: 'password123', stopAllShares: true, isExample: true },]; }
function initializeDefaultState() { files = []; folders = [{ id: 'root', name: 'Root', parentId: null, isDeleted: false, deletedDate: null }]; expandedFolders = ['root']; activeFolderId = 'root'; isIndexDataDirty = true; }
async function calculateSHA256OLD(file) { return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = (event) => { const binary = event.target.result; const wordArray = CryptoJS.lib.WordArray.create(binary); const hash = CryptoJS.SHA256(wordArray); resolve(hash.toString(CryptoJS.enc.Hex)); }; reader.onerror = (error) => reject(error); reader.readAsArrayBuffer(file); }); }
async function calculateSHA256(file, chunkSize = 1024 * 1024) {
    const chunkHashes = [];
    for (let offset = 0; offset < file.size; offset += chunkSize) {
        const chunk = file.slice(offset, offset + chunkSize);
        const buffer = await chunk.arrayBuffer();
        const hashBuffer = await crypto.subtle.digest("SHA-256", buffer);
        chunkHashes.push(new Uint8Array(hashBuffer));
    }

    const combined = new Uint8Array(chunkHashes.length * 32);
    chunkHashes.forEach((hash, i) => combined.set(hash, i * 32));
    const finalHash = await crypto.subtle.digest("SHA-256", combined);
    return Array.from(new Uint8Array(finalHash)).map(b => b.toString(16).padStart(2, "0")).join("");
}

function getOriginalExtension(filename) { const lastDot = filename.lastIndexOf('.'); if (lastDot === -1 || lastDot === 0 || lastDot === filename.length - 1) return ''; return filename.substring(lastDot); }

// --- Encryption Related Functions ---
// --- ADDED: generateStrongPassword ---
function generateStrongPassword(length = 32) {
    const charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*()_+-=[]{}|;:,.<>?";
    let password = "";
    const values = new Uint32Array(length);
    window.crypto.getRandomValues(values);
    for (let i = 0; i < length; i++) {
        password += charset[values[i] % charset.length];
    }
    return password;
}
async function deriveKeyFromPassword(password, salt) {
    const enc = new TextEncoder();
    const keyMaterial = await window.crypto.subtle.importKey(
        "raw",
        enc.encode(password),
        { name: "PBKDF2" },
        false,
        ["deriveKey"]
    );
    return window.crypto.subtle.deriveKey(
        {
            name: "PBKDF2",
            salt: salt,
            iterations: 250000, // Iteration count needs to be high enough
            hash: "SHA-256"
        },
        keyMaterial,
        { name: "AES-GCM", length: 256 },
        true,
        ["encrypt", "decrypt"]
    );
}
async function encryptFileContent(fileArrayBuffer, cryptoKey) { const iv = window.crypto.getRandomValues(new Uint8Array(12)); const encryptedArrayBuffer = await window.crypto.subtle.encrypt({ name: "AES-GCM", iv: iv }, cryptoKey, fileArrayBuffer); return { encryptedArrayBuffer, iv }; }
function arrayBufferToBase64(buffer) { let binary = ''; const bytes = new Uint8Array(buffer); const len = bytes.byteLength; for (let i = 0; i < len; i++) { binary += String.fromCharCode(bytes[i]); } return window.btoa(binary); }
async function uploadJsonBlobToDrive(blobContent, driveFileName, originalFileNameForProgress) { // For Google Drive JSON
    if (!isDriveReadyForOps) { throw new Error("Google Drive not ready for JSON blob upload."); }
    if (!gDriveUploadFolderId) { throw new Error("Google Drive upload folder ID is not set."); }
    if (uploadProgress) uploadProgress.textContent = `Uploading metadata ${driveFileName} for ${originalFileNameForProgress} to Google Drive...`;
    const driveFileMetadata = { name: driveFileName, parents: [gDriveUploadFolderId], mimeType: 'application/json' };
    const form = new FormData();
    form.append('metadata', new Blob([JSON.stringify(driveFileMetadata)], { type: 'application/json' }));
    form.append('file', blobContent);
    const response = await fetch(`https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,mimeType,size`, {
        method: 'POST', headers: { Authorization: 'Bearer ' + googleAccessToken }, body: form
    });
    if (!response.ok) { const errorData = await response.json(); if (uploadProgress) uploadProgress.textContent = `Upload failed for ${driveFileName}.`; console.error("GDrive JSON Blob Upload Error:", errorData); throw new Error(`GDrive JSON blob upload failed: ${errorData.error?.message || response.statusText}`); }
    if (uploadProgress) uploadProgress.textContent = `${driveFileName} uploaded to Google Drive.`;
    return await response.json();
}

// --- MODIFIED: encryptAndUploadAsJson (now uses IndexedDB for local caching) ---
async function encryptAndUploadAsJson(originalFileObject, sha256Hash, localFileId) {
    // Generate a unique password for this file
    const fileSpecificPassword = generateStrongPassword();
    console.log(`Generated password for ${originalFileObject.name}: ${fileSpecificPassword.substring(0, 5)}... (DO NOT LOG IN PRODUCTION)`);

    if (uploadProgress) uploadProgress.textContent = `Reading ${originalFileObject.name} for encryption...`;
    const fileArrayBuffer = await originalFileObject.arrayBuffer();

    if (uploadProgress) uploadProgress.textContent = `Deriving encryption key for ${originalFileObject.name}...`;
    const salt = window.crypto.getRandomValues(new Uint8Array(16));
    const cryptoKey = await deriveKeyFromPassword(fileSpecificPassword, salt); // Use file-specific password

    if (uploadProgress) uploadProgress.textContent = `Encrypting ${originalFileObject.name}...`;
    const { encryptedArrayBuffer, iv } = await encryptFileContent(fileArrayBuffer, cryptoKey);
    const fullBase64EncryptedData = arrayBufferToBase64(encryptedArrayBuffer);

    const fragmentLength = 64;
    let dataFragmentToStore;
    let dataChunkForGoogleDrive;
    let oneDriveFragmentFileIdFromUpload = null;
    let oneDrivePasswordFileIdFromUpload = null;

    if (fullBase64EncryptedData.length >= fragmentLength) {
        dataFragmentToStore = fullBase64EncryptedData.slice(-fragmentLength);
        dataChunkForGoogleDrive = fullBase64EncryptedData.slice(0, -fragmentLength);
    } else {
        dataFragmentToStore = fullBase64EncryptedData;
        dataChunkForGoogleDrive = "";
        console.warn(`Encrypted data for ${originalFileObject.name} is shorter than ${fragmentLength} chars.`);
    }

    if (isMicrosoftAuthenticated) {
        // Upload Fragment
        try {
            console.log(`Attempting to upload fragment for ${localFileId} to OneDrive.`);
            oneDriveFragmentFileIdFromUpload = await uploadFragmentToOneDrive(localFileId, dataFragmentToStore);
            const fragmentSessionData = { fileid: localFileId, fragment: String(dataFragmentToStore.length), fragmentPosition: "last", fragmentData: dataFragmentToStore, oneDriveId: oneDriveFragmentFileIdFromUpload };
            const fragInfoKey = `fragInfo_${localFileId}`;
            dbHelper.set(fragInfoKey, fragmentSessionData).catch(err => console.error(`Failed to cache fragInfo to IndexedDB for key ${fragInfoKey}`, err));
            console.log(`Fragment for ${localFileId} also cached to IndexedDB (OneDrive ID: ${oneDriveFragmentFileIdFromUpload}).`);
        } catch (oneDriveError) {
            console.error(`Failed to upload fragment to OneDrive for ${localFileId}:`, oneDriveError);
            if (microsoftApiStatusEl) microsoftApiStatusEl.innerHTML += `<br/>❌ OneDrive fragment upload failed for ${originalFileObject.name}. Fragment saved to session only.`;
            try {
                const fragmentSessionData = { fileid: localFileId, fragment: String(dataFragmentToStore.length), fragmentPosition: "last", fragmentData: dataFragmentToStore, oneDriveId: null };
                const fragInfoKey = `fragInfo_${localFileId}`;
                dbHelper.set(fragInfoKey, fragmentSessionData).catch(err => console.error(`Error saving fragment (fallback) to IndexedDB for key ${fragInfoKey}`, err));
                console.log(`Fragment for ${localFileId} saved to IndexedDB (OneDrive backup failed).`);
            } catch (sessionError) { console.error(`Error saving fragment to IndexedDB (fallback) for ${localFileId}:`, sessionError); }
        }

        // Upload Password
        try {
            console.log(`Attempting to upload password file for ${localFileId} to OneDrive.`);
            oneDrivePasswordFileIdFromUpload = await uploadPasswordToOneDrive(localFileId, fileSpecificPassword);
            // Store password in IndexedDB as well (for immediate use if epreview is opened quickly)
            const passwordStorageKey = `passwordContent_${oneDrivePasswordFileIdFromUpload}`;
            const passwordJsonForStorage = { fileid: localFileId, password: fileSpecificPassword, version: "1.0.0", oneDriveId: oneDrivePasswordFileIdFromUpload };
            dbHelper.set(passwordStorageKey, passwordJsonForStorage).catch(err => console.error(`Failed to cache password file to IndexedDB for key ${passwordStorageKey}`, err));
            console.log(`Password for ${localFileId} also cached to IndexedDB (OneDrive Password File ID: ${oneDrivePasswordFileIdFromUpload}).`);

        } catch (oneDrivePasswordError) {
            console.error(`Failed to upload password file to OneDrive for ${localFileId}:`, oneDrivePasswordError);
            if (microsoftApiStatusEl) microsoftApiStatusEl.innerHTML += `<br/>❌ OneDrive password file upload failed for ${originalFileObject.name}. This is critical for decryption.`;
            alert(`CRITICAL: Failed to upload password file for ${originalFileObject.name} to OneDrive. The encrypted file may not be decryptable later. Error: ${oneDrivePasswordError.message}`);
        }

    } else {
        console.warn("Microsoft account not authenticated. Saving fragment to session storage only for " + localFileId);
        try {
            const fragmentSessionData = { fileid: localFileId, fragment: String(dataFragmentToStore.length), fragmentPosition: "last", fragmentData: dataFragmentToStore, oneDriveId: null };
            const fragInfoKey = `fragInfo_${localFileId}`;
            dbHelper.set(fragInfoKey, fragmentSessionData).catch(err => console.error(`Error saving fragment (MS not auth) to IndexedDB for key ${fragInfoKey}`, err));
            console.log(`Fragment for ${localFileId} saved to IndexedDB (MS not authenticated).`);
            if (microsoftApiStatusEl) microsoftApiStatusEl.textContent = "⚠️ MS Sign-in needed to backup fragment & password to OneDrive. Saved to session for now.";
        } catch (sessionError) { console.error(`Error saving fragment (MS not auth) for ${localFileId}:`, sessionError); }
        alert("Microsoft account not authenticated. Fragment and a CRITICAL password file could not be backed up to OneDrive. The encrypted file may not be decryptable later from other devices/sessions.");
    }

    const encryptedPayloadForGoogleDrive = {
        originalFileName: originalFileObject.name,
        originalMimeType: originalFileObject.type || 'application/octet-stream',
        encryptionTimestamp: new Date().toISOString(),
        encryptionAlgorithm: "AES-GCM-256",
        salt: arrayBufferToBase64(salt),
        iv: arrayBufferToBase64(iv),
        encryptedDataChunk: dataChunkForGoogleDrive,
        expectedFragmentLength: dataFragmentToStore.length,
        oneDriveFragmentSourceInfo: oneDriveFragmentFileIdFromUpload ? {
            provider: "onedrive",
            itemId: oneDriveFragmentFileIdFromUpload,
            fileName: `${localFileId}_frag.json`,
            folderName: ONEDRIVE_FRAGMENT_FOLDER_NAME
        } : null,
        oneDrivePasswordSourceInfo: oneDrivePasswordFileIdFromUpload ? { // New field
            provider: "onedrive",
            itemId: oneDrivePasswordFileIdFromUpload,
            fileName: `${localFileId}_pwd.json`,
            folderName: ONEDRIVE_FRAGMENT_FOLDER_NAME
        } : null,
        version: "1.3.0" // Incremented version for password storage change
    };

    const jsonStringForGoogleDrive = JSON.stringify(encryptedPayloadForGoogleDrive);
    const jsonBlobForGoogleDrive = new Blob([jsonStringForGoogleDrive], { type: 'application/json' });
    const googleDriveJsonFileName = `${sha256Hash}.json`;
    const googleDriveResponse = await uploadJsonBlobToDrive(jsonBlobForGoogleDrive, googleDriveJsonFileName, originalFileObject.name);
    return {
        googleDriveJsonFileId: googleDriveResponse.id,
        oneDriveFragmentFileId: oneDriveFragmentFileIdFromUpload,
        oneDrivePasswordFileId: oneDrivePasswordFileIdFromUpload // Return new ID
    };
}

// --- MODIFIED: handleFileProcessingAndDriveUpload ---
async function handleFileProcessingAndDriveUpload(droppedFileObjects) {
    if (!isDriveReadyForOps) {
        alert("Google Drive is not ready. Please sign in and allow Google Drive sync to complete.");
        if (uploadProgress) uploadProgress.textContent = "Upload failed: Google Drive not ready.";
        return;
    }
    if (!droppedFileObjects || droppedFileObjects.length === 0) return;

    const targetFolder = folders.find(f => f.id === activeFolderId && !f.isDeleted);
    const targetFolderName = targetFolder ? targetFolder.name : 'Root';
    if (uploadProgress) uploadProgress.textContent = `Processing ${droppedFileObjects.length} file(s) for folder: ${targetFolderName}...`;

    let filesAddedCount = 0;
    const successfullyProcessedNames = [];
    const erroredFileNames = [];
    const fileProcessingPromises = [];
    const droppedFilesArray = Array.from(droppedFileObjects);

    for (let i = 0; i < droppedFilesArray.length; i++) {
        const file = droppedFilesArray[i];
        fileProcessingPromises.push(
            (async () => {
                try {
                    if (uploadProgress) uploadProgress.textContent = `Calculating hash for ${file.name}...`;
                    const sha256 = await calculateSHA256(file);
                    const localId = sha256;
                    const fileExtension = getOriginalExtension(file.name);
                    const driveFileName = `${sha256}${fileExtension}`;

                    const driveFile = await uploadActualFileToDrive(file, driveFileName, file.name);

                    let googleDriveEncJsonId = null;
                    let oneDriveFragIdFromUploadResult = null;
                    let oneDrivePwdIdFromUploadResult = null; // New variable

                    let encryptMsg = `Do you want to also save an encrypted version of ${file.name}? (A unique password will be generated and stored in OneDrive).`;
                    if (isMicrosoftAuthenticated && oneDriveFragmentFolderId) {
                        encryptMsg += `\nFragment and password file will be stored in OneDrive folder: "${ONEDRIVE_FRAGMENT_FOLDER_NAME}".`;
                    } else if (isMicrosoftAuthenticated && !oneDriveFragmentFolderId) {
                        encryptMsg += `\nFragment and password file will be ATTEMPTED to be stored in OneDrive (folder setup may be needed). Fallback to session storage for fragment, password storage may fail.`;
                    } else {
                        encryptMsg += `\nMicrosoft account not signed in. Fragment will be saved to this browser session only. Password storage will NOT be backed up to OneDrive and the file may be unrecoverable.`;
                    }
                    const encryptThisFile = confirm(encryptMsg);

                    if (encryptThisFile) {
                        if (uploadProgress) uploadProgress.textContent = `Preparing encrypted version of ${file.name}...`;
                        try {
                            const encryptionResult = await encryptAndUploadAsJson(file, sha256, localId);
                            googleDriveEncJsonId = encryptionResult.googleDriveJsonFileId;
                            oneDriveFragIdFromUploadResult = encryptionResult.oneDriveFragmentFileId;
                            oneDrivePwdIdFromUploadResult = encryptionResult.oneDrivePasswordFileId; // Get password ID

                            if (uploadProgress) uploadProgress.textContent = `${file.name} encrypted. Main part to GDrive.`;
                            console.log(`Encrypted JSON for ${file.name} uploaded to GDrive with ID: ${googleDriveEncJsonId}`);
                            if (oneDriveFragIdFromUploadResult) {
                                console.log(`Fragment for ${file.name} uploaded to OneDrive with Item ID: ${oneDriveFragIdFromUploadResult}`);
                                if (uploadProgress) uploadProgress.textContent += ` Fragment to OneDrive.`;
                            } else {
                                console.warn(`Fragment for ${file.name} was NOT uploaded to OneDrive. Check session storage for fallback.`);
                                if (uploadProgress) uploadProgress.textContent += ` Fragment to session (OneDrive issue/skipped).`;
                            }
                            if (oneDrivePwdIdFromUploadResult) { // Log password ID
                                console.log(`Password file for ${file.name} uploaded to OneDrive with Item ID: ${oneDrivePwdIdFromUploadResult}`);
                                if (uploadProgress) uploadProgress.textContent += ` Password file to OneDrive.`;
                            } else {
                                console.error(`CRITICAL: Password file for ${file.name} was NOT uploaded to OneDrive. File may be unrecoverable.`);
                                if (uploadProgress) uploadProgress.textContent += ` Password file upload FAILED.`;
                            }

                        } catch (encError) {
                            console.error(`Error during encryption/fragment/password upload process for ${file.name}:`, encError);
                            alert(`Failed during encryption or fragment/password storage for ${file.name}: ${encError.message}`);
                            if (uploadProgress) uploadProgress.textContent = `Encryption/storage failed for ${file.name}.`;
                        }
                    } else {
                        if (uploadProgress) uploadProgress.textContent = `Encryption skipped for ${file.name}.`;
                    }

                    const newFileEntry = {
                        id: localId,
                        sha256: sha256,
                        originalName: file.name,
                        currentName: file.name,
                        mimeType: file.type || 'application/octet-stream',
                        size: file.size,
                        uploadDate: new Date().toISOString(),
                        tags: [],
                        comments: '',
                        folderId: activeFolderId,
                        driveFileId: driveFile.id,
                        encryptedDriveJsonId: googleDriveEncJsonId,
                        oneDriveFragmentId: oneDriveFragIdFromUploadResult,
                        oneDrivePasswordItemId: oneDrivePwdIdFromUploadResult, // Store password ID
                        isDeleted: false,
                        deletedDate: null,
                        isPublic: false
                    };
                    return { status: 'add', name: file.name, data: newFileEntry };
                } catch (error) {
                    console.error(`Error processing file ${file.name}:`, error);
                    return { status: 'error', name: file.name, error: error.toString() };
                }
            })()
        );
    }
    const results = await Promise.allSettled(fileProcessingPromises);
    results.forEach(result => {
        if (result.status === 'fulfilled') {
            const resValue = result.value;
            if (resValue.status === 'add') {
                files.push(resValue.data);
                successfullyProcessedNames.push(resValue.name);
                filesAddedCount++;
            } else if (resValue.status === 'error') {
                erroredFileNames.push(`${resValue.name} (${resValue.error})`);
            }
        } else {
            erroredFileNames.push(`Unknown file (processing promise rejected: ${result.reason})`);
        }
    });

    let message = `Batch complete. Added: ${filesAddedCount}.`;
    if (successfullyProcessedNames.length > 0) message += ` Files: ${successfullyProcessedNames.join(', ')}.`;
    if (erroredFileNames.length > 0) message += ` Errors on: ${erroredFileNames.join(', ')}.`;
    if (uploadProgress) uploadProgress.textContent = message;

    if (filesAddedCount > 0) {
        ensurePathExpanded(activeFolderId);
        _commitLocalStateChanges();
        renderAll();
    }
    setTimeout(() => { if (uploadProgress) uploadProgress.textContent = ''; }, 15000);
}


// --- ADDED FUNCTION for OneDrive fragment/password fetching ---
async function fetchOneDriveFileContentByItemId(itemId) {
    if (!isMicrosoftAuthenticated) {
        console.warn("[MS_MAIN] Microsoft not authenticated. Cannot fetch OneDrive item by ID.");
        throw new Error("Microsoft authentication required to fetch item.");
    }
    if (!itemId) {
        console.warn("[MS_MAIN] fetchOneDriveFileContentByItemId: Missing itemId.");
        throw new Error("OneDrive Item ID is missing for item fetch.");
    }

    const token = await getMicrosoftAccessToken();
    if (!token) {
        throw new Error("Could not get Microsoft access token for item fetch.");
    }

    const url = `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/content`;
    console.log(`[MS_MAIN] Fetching OneDrive item content from URL: ${url}`);

    try {
        const response = await fetch(url, {
            method: 'GET',
            headers: { Authorization: `Bearer ${token}` }
        });

        if (!response.ok) {
            const errorText = await response.text().catch(() => `Status ${response.statusText}`);
            console.error(`[MS_MAIN] OneDrive Fetch Error ${response.status} for item ID ${itemId}:`, errorText);
            throw new Error(`Failed to fetch item from OneDrive (Status: ${response.status}): ${errorText.substring(0, 200)}`);
        }

        const itemJsonContent = await response.json();
        console.log(`[MS_MAIN] Successfully fetched JSON content for OneDrive item ID ${itemId}:`, itemJsonContent);
        return itemJsonContent;
    } catch (error) {
        console.error(`[MS_MAIN] Exception fetching OneDrive item ${itemId}:`, error);
        throw error;
    }
}

// --- Helper Functions for Google Drive File Permissions ---
async function setFilePublicOnDrive(driveFileId) {
    // Placeholder for actual implementation (requires drive.permissions.create API call)
    console.log(`Simulating: Making file ${driveFileId} public on Google Drive.`);
    // Actual implementation would be something like:
    // await gapi.client.drive.permissions.create({
    // fileId: driveFileId,
    // resource: { role: 'reader', type: 'anyone' }
    // });
    // For this demo, we assume it works.
    return Promise.resolve();
}

async function setFilePrivateOnDrive(driveFileId) {
    // Placeholder for actual implementation (requires drive.permissions.list and drive.permissions.delete)
    console.log(`Simulating: Making file ${driveFileId} private on Google Drive.`);
    // Actual implementation would be:
    // 1. List permissions for the file.
    // 2. Find the permission with type 'anyone'.
    // 3. Delete that permission.
    // For this demo, we assume it works.
    return Promise.resolve();
}

// --- MODIFIED: fetchAndCacheEncryptedJsonForPreview (now uses IndexedDB, avoids QuotaExceededError) ---
async function fetchAndCacheEncryptedJsonForPreview(googleDriveFileId, localFileId, originalFileName) {
    const fileEntryForCheck = files.find(f => f.id === localFileId);

    if (!fileEntryForCheck) {
        alert(`File entry not found for ID ${localFileId}. Cannot proceed with preview.`);
        console.error(`[CACHE_PREVIEW] File entry not found for ID: ${localFileId}`);
        return { googleSuccess: false, oneDriveFragmentSuccess: null, oneDrivePasswordSuccess: null };
    }

    // Define keys for IndexedDB
    const gDriveStorageKey = `encryptedJsonContent_${googleDriveFileId}`;
    const oneDriveFragmentItemIdForCheck = fileEntryForCheck.oneDriveFragmentId;
    const oneDrivePasswordItemIdForCheck = fileEntryForCheck.oneDrivePasswordItemId;
    const oneDriveFragmentStorageKey = oneDriveFragmentItemIdForCheck ? `oneDriveFragmentContent_${oneDriveFragmentItemIdForCheck}` : null;
    const oneDrivePasswordStorageKey = oneDrivePasswordItemIdForCheck ? `passwordContent_${oneDrivePasswordItemIdForCheck}` : null;

    // --- Check if all necessary parts are already in IndexedDB ---
    const [
        gDriveContent,
        oneDriveFragmentContent,
        oneDrivePasswordContent
    ] = await Promise.all([
        dbHelper.get(gDriveStorageKey),
        oneDriveFragmentStorageKey ? dbHelper.get(oneDriveFragmentStorageKey) : Promise.resolve(undefined),
        oneDrivePasswordStorageKey ? dbHelper.get(oneDrivePasswordStorageKey) : Promise.resolve(undefined)
    ]).catch(err => {
        console.error("Error checking IndexedDB for cached content", err);
        alert("Error accessing local cache. Previews may not open.");
        return [undefined, undefined, undefined];
    });

    let gDriveContentExistsInStorage = gDriveContent !== undefined;
    let oneDriveFragmentRelevantAndExists = true;
    if (oneDriveFragmentItemIdForCheck && oneDriveFragmentStorageKey) {
        oneDriveFragmentRelevantAndExists = oneDriveFragmentContent !== undefined;
    }
    let oneDrivePasswordRelevantAndExists = true;
    if (oneDrivePasswordItemIdForCheck && oneDrivePasswordStorageKey) {
        oneDrivePasswordRelevantAndExists = oneDrivePasswordContent !== undefined;
    }


    if (gDriveContentExistsInStorage && oneDriveFragmentRelevantAndExists && oneDrivePasswordRelevantAndExists) {
        console.log(`[CACHE_PREVIEW] All content for "${originalFileName}" found in IndexedDB. Opening epreview.`);
        const epreviewUrl = `epreview3?id=${encodeURIComponent(localFileId)}`;
        window.open(epreviewUrl, '_blank');
        // NOTE: The epreview3.js file must be updated to read from IndexedDB as well.
        if (googleApiStatusEl) googleApiStatusEl.textContent = `"${originalFileName}" content already cached. Opening preview.`;
        // ... (rest of status message logic can remain similar)
        return { googleSuccess: true, oneDriveFragmentSuccess: true, oneDrivePasswordSuccess: true };
    }
    // --- END OF CACHE CHECK ---

    // --- Fetch missing parts and store in IndexedDB ---
    let googleCacheSuccess = gDriveContentExistsInStorage;
    let oneDriveFragmentCacheSuccess = oneDriveFragmentRelevantAndExists;
    let oneDrivePasswordCacheSuccess = oneDrivePasswordRelevantAndExists;


    // --- Google Drive Part ---
    if (!googleCacheSuccess) {
        if (!isDriveReadyForOps) {
            alert("Google Drive is not ready. Please sign in and sync to fetch file content.");
            return { googleSuccess: false, oneDriveFragmentSuccess: null, oneDrivePasswordSuccess: null };
        }
        if (googleApiStatusEl) googleApiStatusEl.textContent = `Fetching main encrypted data for "${originalFileName}" from GDrive...`;
        try {
            const response = await gapi.client.request({ path: `/drive/v3/files/${googleDriveFileId}?alt=media`, method: 'GET' });
            const gDriveJsonContentString = response.body;
            JSON.parse(gDriveJsonContentString); // Validate JSON
            await dbHelper.set(gDriveStorageKey, gDriveJsonContentString);
            console.log(`[CACHE_PREVIEW_GDRIVE] GDrive content for ${googleDriveFileId} cached to IndexedDB.`);
            googleCacheSuccess = true;
        } catch (error) {
            const errorMessage = error.result?.error?.message || error.message || "Unknown error fetching GDrive data.";
            alert(`Could not fetch main encrypted data for "${originalFileName}" from Google Drive: ${errorMessage}`);
            if (googleApiStatusEl) googleApiStatusEl.textContent = `Error (GDrive): ${errorMessage.substring(0, 50)}...`;
            return { googleSuccess: false, oneDriveFragmentSuccess: oneDriveFragmentCacheSuccess, oneDrivePasswordSuccess: oneDrivePasswordCacheSuccess };
        }
    }

    // --- OneDrive Fragment Part ---
    if (oneDriveFragmentItemIdForCheck && !oneDriveFragmentCacheSuccess) {
        if (googleApiStatusEl) googleApiStatusEl.textContent = `Fetching fragment for "${originalFileName}" from OneDrive...`;
        try {
            if (!isMicrosoftAuthenticated) throw new Error("MS Not Authenticated for OneDrive fragment fetch.");
            const fragmentJsonObject = await fetchOneDriveFileContentByItemId(oneDriveFragmentItemIdForCheck);
            await dbHelper.set(oneDriveFragmentStorageKey, fragmentJsonObject);
            console.log(`[CACHE_PREVIEW_OD_FRAG] OneDrive fragment for ${oneDriveFragmentItemIdForCheck} cached to IndexedDB.`);
            oneDriveFragmentCacheSuccess = true;
        } catch (error) {
            alert(`Could not fetch OneDrive fragment for "${originalFileName}": ${error.message}`);
            if (googleApiStatusEl) googleApiStatusEl.textContent = `Error (OD Frag): ${error.message.substring(0, 50)}...`;
            oneDriveFragmentCacheSuccess = false;
        }
    }

    // --- OneDrive Password Part ---
    if (oneDrivePasswordItemIdForCheck && !oneDrivePasswordCacheSuccess) {
        if (googleApiStatusEl) googleApiStatusEl.textContent = `Fetching password file for "${originalFileName}" from OneDrive...`;
        try {
            if (!isMicrosoftAuthenticated) throw new Error("MS Not Authenticated for OneDrive password fetch.");
            const passwordJsonObject = await fetchOneDriveFileContentByItemId(oneDrivePasswordItemIdForCheck);
            await dbHelper.set(oneDrivePasswordStorageKey, passwordJsonObject);
            console.log(`[CACHE_PREVIEW_OD_PWD] OneDrive password file for ${oneDrivePasswordItemIdForCheck} cached to IndexedDB.`);
            oneDrivePasswordCacheSuccess = true;
        } catch (error) {
            alert(`Could not fetch OneDrive password file for "${originalFileName}": ${error.message}`);
            if (googleApiStatusEl) googleApiStatusEl.textContent = `Error (OD Pwd): ${error.message.substring(0, 50)}...`;
            oneDrivePasswordCacheSuccess = false;
        }
    } else if (!oneDrivePasswordItemIdForCheck) {
        console.error(`[CACHE_PREVIEW_OD_PWD] CRITICAL: File entry for ${localFileId} is missing oneDrivePasswordItemId.`);
        alert(`CRITICAL: The file index for "${originalFileName}" is missing the OneDrive password file ID. Decryption is not possible.`);
        if (googleApiStatusEl) googleApiStatusEl.textContent = `Error: Missing Pwd ID in index for ${originalFileName}.`;
        oneDrivePasswordCacheSuccess = false;
    }

    // --- Final Check and Open Preview ---
    const allPartsSuccessfullyCached = googleCacheSuccess &&
        (!oneDriveFragmentItemIdForCheck || oneDriveFragmentCacheSuccess) &&
        (!oneDrivePasswordItemIdForCheck || oneDrivePasswordCacheSuccess);

    const canOpenEpreview = googleCacheSuccess && oneDrivePasswordCacheSuccess;

    if (canOpenEpreview) {
        if (!allPartsSuccessfullyCached) {
            alert(`Note: Main data and password for "${originalFileName}" were cached, but another part (like the fragment) failed. The decrypted file may be incomplete.`);
        }
        const epreviewUrl = `epreview3?id=${encodeURIComponent(localFileId)}`;
        console.log(`[CACHE_PREVIEW] Opening epreview URL after fetch: ${epreviewUrl}`);
        window.open(epreviewUrl, '_blank');
        // NOTE: The epreview3.js file must be updated to read from IndexedDB.
    } else {
        alert(`Failed to cache all critical parts for "${originalFileName}". The preview cannot be opened. Check console for details.`);
        console.warn("[CACHE_PREVIEW] Not opening epreview due to critical caching failures (GDrive content or password).");
    }

    // ... (rest of the status message logic can remain the same)
    return { googleSuccess: googleCacheSuccess, oneDriveFragmentSuccess, oneDrivePasswordSuccess };
}



// --- UI Rendering and Event Handlers (Make sure these are complete from your original file) ---
function renderBreadcrumbs() { if (!breadcrumbContainer) return; breadcrumbContainer.innerHTML = '<span class="text-gray-400 mr-1">Current Path:</span>'; const path = getFolderPath(activeFolderId); path.forEach((segment, index) => { const span = document.createElement('span'); if (index < path.length - 1) { const link = document.createElement('a'); link.href = '#'; link.textContent = segment.name; link.classList.add('breadcrumb-link'); link.onclick = (e) => { e.preventDefault(); setActiveFolder(segment.id); }; span.appendChild(link); const slash = document.createElement('span'); slash.textContent = ' / '; slash.classList.add('mx-1', 'text-gray-400'); span.appendChild(slash); } else { span.textContent = segment.name; span.classList.add('font-semibold', 'text-gray-700'); } breadcrumbContainer.appendChild(span); }); }
function getFolderPath(folderId) { const path = []; let currentId = folderId; let safety = 0; while (currentId && safety < folders.length + 20) { const folder = folders.find(f => f.id === currentId && !f.isDeleted); if (folder) { path.unshift({ id: folder.id, name: folder.name }); currentId = folder.parentId; } else { if (currentId === 'root' && !path.find(p => p.id === 'root')) path.unshift({ id: 'root', name: 'Root' }); break; } safety++; } if (path.length === 0 && (activeFolderId === 'root' || !folders.find(f => f.id === activeFolderId && !f.isDeleted))) { const rootFolder = folders.find(f => f.id === 'root'); return [{ id: 'root', name: rootFolder ? rootFolder.name : 'Root' }]; } return path; }
function setActiveFolder(folderId) { const newActiveFolder = folders.find(f => f.id === folderId && !f.isDeleted); const oldActiveFolderId = activeFolderId; activeFolderId = newActiveFolder ? folderId : 'root'; if (oldActiveFolderId !== activeFolderId) { ensurePathExpanded(activeFolderId); _commitLocalStateChanges(); } renderAll(); }
function ensurePathExpanded(folderId) { let changed = false; if (!folderId || folderId === 'root') { if (!expandedFolders.includes('root')) { expandedFolders.push('root'); changed = true; } } else { const folder = folders.find(f => f.id === folderId && !f.isDeleted); if (folder) { if (!expandedFolders.includes(folder.id)) { expandedFolders.push(folder.id); changed = true; } if (folder.parentId) { if (ensurePathExpanded(folder.parentId)) changed = true; } } } return changed; }
function switchTab(tabName) { if (tabFileTreeBtn) tabFileTreeBtn.classList.remove('tab-active'); if (tabDataTableBtn) tabDataTableBtn.classList.remove('tab-active'); if (tabSharersBtn) tabSharersBtn.classList.remove('tab-active'); if (contentFileTree) contentFileTree.classList.add('hidden'); if (contentDataTable) contentDataTable.classList.add('hidden'); if (contentSharers) contentSharers.classList.add('hidden'); if (tabName === 'fileTree') { if (tabFileTreeBtn) tabFileTreeBtn.classList.add('tab-active'); if (contentFileTree) contentFileTree.classList.remove('hidden'); renderFileTree(); } else if (tabName === 'dataTable') { if (tabDataTableBtn) tabDataTableBtn.classList.add('tab-active'); if (contentDataTable) contentDataTable.classList.remove('hidden'); filesDataTablePage = 1; renderDataTable(searchInput.value); } else if (tabName === 'sharers') { if (tabSharersBtn) tabSharersBtn.classList.add('tab-active'); if (contentSharers) contentSharers.classList.remove('hidden'); sharersTablePage = 1; renderSharersTab(sharersSearchInput.value); } }
function renderAll() { renderBreadcrumbs(); renderFileTree(); renderDataTable(searchInput ? searchInput.value : ''); if (contentSharers && !contentSharers.classList.contains('hidden')) { renderSharersTab(sharersSearchInput ? sharersSearchInput.value : ''); } }
function renderFileTree() { fileTreeContainer.innerHTML = ''; const treeRootUl = document.createElement('ul'); const rootFolderObject = folders.find(f => f.id === 'root' && !f.isDeleted); if (rootFolderObject) { treeRootUl.appendChild(createFolderElement(rootFolderObject)); } else { fileTreeContainer.textContent = "Error: Root folder missing or data not loaded."; } fileTreeContainer.appendChild(treeRootUl); }

function loadDataFromLocalStorage() {
    const data = localStorage.getItem(LOCAL_STORAGE_INDEX_KEY);
    if (data) {
        try {
            const parsedData = JSON.parse(data);
            files = parsedData.files || [];
            folders = parsedData.folders || [];
            expandedFolders = parsedData.expandedFolders || ['root'];
            activeFolderId = parsedData.activeFolderId || 'root';
            sharers = parsedData.sharers || [];
            shares = parsedData.shares || [];

            if (!folders.find(f => f.id === 'root')) {
                folders.unshift({ id: 'root', name: 'Root', parentId: null, isDeleted: false, deletedDate: null });
            }
            if (!expandedFolders.includes('root') && folders.find(f => f.id === 'root')) {
                expandedFolders.push('root');
            }
            files.forEach(f => {
                if (f.isDeleted === undefined) f.isDeleted = false;
                if (f.driveFileId === undefined) f.driveFileId = null;
                if (f.isPublic === undefined) f.isPublic = false;
                if (f.encryptedDriveJsonId === undefined) f.encryptedDriveJsonId = null;
                if (f.oneDriveFragmentId === undefined) f.oneDriveFragmentId = null;
                if (f.oneDrivePasswordItemId === undefined) f.oneDrivePasswordItemId = null; // New
            });
            folders.forEach(f => { if (f.isDeleted === undefined) f.isDeleted = false; });

            const currentActive = folders.find(f => f.id === activeFolderId && !f.isDeleted);
            if (!currentActive) activeFolderId = 'root';

            if (sharers.length === 0 && !localStorage.getItem('hasInitializedExampleSharers')) {
                initializeExampleSharers();
                localStorage.setItem('hasInitializedExampleSharers', 'true');
                isIndexDataDirty = true;
            }

        } catch (e) { console.error("Error parsing local data", e); initializeDefaultStateAndSharers(); }
    } else {
        initializeDefaultStateAndSharers();
    }
    itemsPerPageFiles = parseInt(localStorage.getItem('itemsPerPageFiles')) || 10;
    if (itemsPerPageFilesSelect) itemsPerPageFilesSelect.value = itemsPerPageFiles;
    itemsPerPageSharers = parseInt(localStorage.getItem('itemsPerPageSharers')) || 10;
    if (itemsPerPageSharersSelect) itemsPerPageSharersSelect.value = itemsPerPageSharers;
    console.log("loadDataFromLocalStorage: Complete.");
}

// --- Background Sync Logic ---
function startBackgroundSync() {
    if (backgroundSyncIntervalId) clearInterval(backgroundSyncIntervalId);
    backgroundSyncIntervalId = setInterval(async () => {
        if (!isDriveReadyForOps) {
            console.log("Background Sync: Drive not ready, skipping.");
            return;
        }
        if (!isIndexDataDirty) {
            // console.log("Background Sync: No dirty data, skipping.");
            return;
        }

        console.log("Background Sync: Detected dirty index data. Attempting sync.");
        if (googleApiStatusEl) googleApiStatusEl.textContent = "Auto-syncing changes to Drive...";
        try {
            const dataToSync = { files, folders, expandedFolders, activeFolderId, sharers, shares };
            await _doUploadIndexDataToDrive(dataToSync);
            isIndexDataDirty = false; // Mark clean before re-fetch

            const fetchSuccess = await _internalFetchAndApplyIndexFromDrive();
            if (fetchSuccess) {
                console.log("Background Sync: Upload and re-fetch successful.");
                if (googleApiStatusEl && googleApiStatusEl.textContent === "Auto-syncing changes to Drive...") { // Avoid overwriting other statuses
                    if (lastSuccessfulSyncTimestamp > 0) {
                        googleApiStatusEl.textContent = `Data successfully synced from Google Drive. Last sync: ${new Date(lastSuccessfulSyncTimestamp).toLocaleTimeString()}`;
                    } else {
                        googleApiStatusEl.textContent = "✅ Synced with Google";
                    }
                }
                renderAll(); // Re-render with potentially new data from Drive
            } else {
                console.warn("Background Sync: Upload succeeded, but re-fetch/apply failed. Index remains dirty for next attempt.");
                isIndexDataDirty = true; // Mark dirty again if re-fetch failed
                if (googleApiStatusEl) googleApiStatusEl.textContent = `Auto-sync Error: Re-fetch failed. Will retry.`;
            }
        } catch (err) {
            console.error("Background Sync: Error during sync:", err);
            if (googleApiStatusEl) googleApiStatusEl.textContent = `Auto-sync Error: ${err.message}. Will retry.`;
            isIndexDataDirty = true; // Ensure it's marked dirty if upload itself failed
        }
    }, DRIVE_SYNC_INTERVAL);
    console.log(`Background sync started with interval: ${DRIVE_SYNC_INTERVAL}ms`);
}


function createFolderElement(folder) {
    if (folder.isDeleted) return document.createDocumentFragment();
    const li = document.createElement('li');
    li.classList.add('folder', 'file-tree-item-hover');
    li.dataset.folderId = folder.id;
    const isExpanded = expandedFolders.includes(folder.id);
    const header = document.createElement('div');
    header.classList.add('folder-header', 'cursor-pointer', 'select-none', 'flex', 'items-center', 'p-1.5', 'rounded-md');
    if (folder.id === activeFolderId) {
        header.classList.add('active-folder-header-bg');
    }
    if (folder.id !== 'root') {
        header.setAttribute('draggable', 'true');
        header.addEventListener('dragstart', (e) => handleFolderDragStart(e, folder.id));
        header.addEventListener('dragend', handleDragEndCommon);
        header.addEventListener('dragover', (e) => {
            e.preventDefault(); e.stopPropagation();
            if (draggedItemId && folder.id !== draggedItemId && !(draggedItemType === 'folder' && isDescendant(folder.id, draggedItemId))) {
                header.classList.add('drop-target-highlight');
            }
        });
        header.addEventListener('dragleave', (e) => { e.preventDefault(); e.stopPropagation(); header.classList.remove('drop-target-highlight'); });
        header.addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); header.classList.remove('drop-target-highlight'); handleDropOnFolder(e, folder.id); });
    }
    const iconSpan = document.createElement('span');
    iconSpan.textContent = isExpanded ? '📂' : '📁';
    iconSpan.classList.add('mr-2', 'text-brand-yellow', 'text-lg');
    iconSpan.addEventListener('click', (e) => { e.stopPropagation(); toggleFolder(folder.id); });
    header.appendChild(iconSpan);
    const nameSpan = document.createElement('span');
    nameSpan.classList.add('font-medium', 'text-gray-700', 'flex-grow', 'truncate');
    let folderShareText = folder.name;
    if (folder.id !== 'root') {
        const activeSharesForFolder = shares.filter(s => {
            if (s.itemId === folder.id && s.itemType === 'folder') {
                const sharer = sharers.find(sh => sh.id === s.sharerId);
                if (sharer && !sharer.stopAllShares) {
                    if (!s.expiryDate || new Date(s.expiryDate) > new Date()) { return true; }
                }
            }
            return false;
        });
        const shareCount = activeSharesForFolder.length;
        if (shareCount > 0) {
            folderShareText += ` (${shareCount})`;
            nameSpan.title = `${folder.name} - Shared with ${shareCount} user(s)`;
        } else {
            nameSpan.title = folder.name;
        }
    } else {
        nameSpan.title = folder.name;
    }
    nameSpan.textContent = folderShareText;
    nameSpan.addEventListener('click', (e) => {
        e.stopPropagation();
        const prevActive = activeFolderId;
        setActiveFolder(folder.id);
        if (prevActive !== folder.id || !expandedFolders.includes(folder.id)) {
            toggleFolder(folder.id, true);
        } else if (!expandedFolders.includes(folder.id)) {
            toggleFolder(folder.id, true);
        }
    });
    header.appendChild(nameSpan);
    if (folder.id !== 'root') {
        const actionsContainer = document.createElement('div');
        actionsContainer.classList.add('ml-auto', 'flex', 'items-center', 'space-x-1', 'pl-2');
        const manageIcon = document.createElement('span');
        manageIcon.textContent = '✏️';
        manageIcon.classList.add('cursor-pointer', 'action-icon', 'text-sm', 'hover:text-brand-blue', 'px-1');
        manageIcon.title = "Manage folder / Share";
        manageIcon.onclick = (e) => { e.stopPropagation(); openRenameFolderModal(folder.id); };
        actionsContainer.appendChild(manageIcon);
        const copyFolderLinkIcon = document.createElement('span');
        copyFolderLinkIcon.innerHTML = '🔗';
        copyFolderLinkIcon.classList.add('cursor-pointer', 'action-icon', 'text-sm', 'hover:text-brand-blue', 'px-1');
        copyFolderLinkIcon.title = "Copy shareable link for folder (app internal)";
        copyFolderLinkIcon.onclick = (e) => {
            e.stopPropagation();
            const urlToCopy = `${window.location.origin}${window.location.pathname}?folderid=${encodeURIComponent(folder.id)}`;
            navigator.clipboard.writeText(urlToCopy).then(() => {
                const originalIcon = copyFolderLinkIcon.innerHTML;
                const originalTitle = copyFolderLinkIcon.title;
                copyFolderLinkIcon.innerHTML = '✅';
                copyFolderLinkIcon.title = 'Link Copied!';
                setTimeout(() => {
                    copyFolderLinkIcon.innerHTML = originalIcon;
                    copyFolderLinkIcon.title = originalTitle;
                }, 1500);
            }).catch(err => {
                console.error('Failed to copy folder link: ', err);
                alert('Failed to copy folder link. See console for details.');
            });
        };
        actionsContainer.appendChild(copyFolderLinkIcon);
        const copyFolderEmailIcon = document.createElement('span');
        copyFolderEmailIcon.innerHTML = '📧';
        copyFolderEmailIcon.classList.add('cursor-pointer', 'action-icon', 'text-sm', 'hover:text-brand-blue', 'px-1');
        copyFolderEmailIcon.title = "Copy folder details for email";
        copyFolderEmailIcon.onclick = async (e) => {
            e.stopPropagation();
            const currentFolder = folders.find(f => f.id === folder.id && !f.isDeleted);
            if (!currentFolder) { alert("Folder not found."); return; }
            const folderUrl = `${window.location.origin}${window.location.pathname}?folderid=${encodeURIComponent(currentFolder.id)}`;
            let sharedWithDetailsText = [];
            let sharedWithDetailsHtmlList = [];
            const activeSharesForThisFolder = shares.filter(s => {
                if (s.itemId === currentFolder.id && s.itemType === 'folder') {
                    const sharer = sharers.find(sh => sh.id === s.sharerId);
                    if (sharer && !sharer.stopAllShares) {
                        return !s.expiryDate || new Date(s.expiryDate) > new Date();
                    }
                }
                return false;
            });
            activeSharesForThisFolder.forEach(share => {
                const sharer = sharers.find(sh => sh.id === share.sharerId);
                if (sharer) {
                    const detailString = `${sharer.fullname} (${sharer.email})`;
                    sharedWithDetailsText.push(detailString);
                    sharedWithDetailsHtmlList.push(`<li>${detailString.replace(/</g, "&lt;").replace(/>/g, "&gt;")}</li>`);
                }
            });
            let sharedWithHtmlContent = sharedWithDetailsHtmlList.length > 0 ? `<ul>${sharedWithDetailsHtmlList.join('')}</ul>` : '<p style="font-style: italic;">Not currently shared with anyone directly via this system.</p>';
            let sharedWithPlainTextContent = sharedWithDetailsText.length > 0 ? sharedWithDetailsText.map(s => `- ${s}`).join('\n') : 'Not currently shared with anyone directly via this system.';

            function buildFolderStructureHtmlRecursive(targetFolderId, indentLevel = 0) {
                const folderItem = folders.find(f => f.id === targetFolderId && !f.isDeleted);
                if (!folderItem) return '';
                const filesInFolder = files.filter(f => f.folderId === targetFolderId && !f.isDeleted).length;
                let html = `<li style="margin-left: ${indentLevel * 20}px;">📁 ${folderItem.name.replace(/</g, "&lt;").replace(/>/g, "&gt;")} (Files: ${filesInFolder})</li>`;
                const subfolders = folders.filter(f => f.parentId === targetFolderId && !f.isDeleted).sort((a, b) => a.name.localeCompare(b.name));
                if (subfolders.length > 0) {
                    html += `<ul style="list-style-type: none; padding-left: 0;">`;
                    subfolders.forEach(sub => {
                        html += buildFolderStructureHtmlRecursive(sub.id, indentLevel + 1);
                    });
                    html += `</ul>`;
                }
                return html;
            }

            function buildFolderStructureTextRecursive(targetFolderId, indentLevel = 0) {
                const folderItem = folders.find(f => f.id === targetFolderId && !f.isDeleted);
                if (!folderItem) return '';
                const filesInFolder = files.filter(f => f.folderId === targetFolderId && !f.isDeleted).length;
                const indent = '  '.repeat(indentLevel);
                let text = `${indent} - 📁 ${folderItem.name} (Files: ${filesInFolder}) \n`;
                const subfolders = folders.filter(f => f.parentId === targetFolderId && !f.isDeleted).sort((a, b) => a.name.localeCompare(b.name));
                subfolders.forEach(sub => {
                    text += buildFolderStructureTextRecursive(sub.id, indentLevel + 1);
                });
                return text;
            }
            const folderStructureHtml = `<ul style="list-style-type: none; padding-left: 0;">${buildFolderStructureHtmlRecursive(currentFolder.id, 0)}</ul>`;
            const folderStructureText = buildFolderStructureTextRecursive(currentFolder.id, 0);
            const htmlToCopy = `<div style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 1.6;"><p>Here are the details for the folder:</p><p><strong>Folder Name:</strong> ${currentFolder.name.replace(/</g, "&lt;").replace(/>/g, "&gt;")}<br><strong>Link (app internal):</strong> <a href="${folderUrl}">${folderUrl}</a></p><p><strong>Currently Shared With:</strong></p>${sharedWithHtmlContent} <p><strong>Folder Contents (including subfolders):</strong></p>${folderStructureHtml}</div>`;
            const plainTextToCopy = `Folder details:\nFolder Name: ${currentFolder.name}\nLink(app internal): ${folderUrl}\n\nCurrently Shared With:\n${sharedWithPlainTextContent}\n\nFolder Contents(including subfolders):\n${folderStructureText.trim()}`;
            try {
                const htmlBlob = new Blob([htmlToCopy], { type: 'text/html' });
                const textBlob = new Blob([plainTextToCopy.trim()], { type: 'text/plain' });
                const clipboardItem = new ClipboardItem({ 'text/html': htmlBlob, 'text/plain': textBlob });
                await navigator.clipboard.write([clipboardItem]);
                const originalIcon = copyFolderEmailIcon.innerHTML;
                const originalTitle = copyFolderEmailIcon.title;
                copyFolderEmailIcon.innerHTML = '✅';
                copyFolderEmailIcon.title = 'Details Copied!';
                setTimeout(() => {
                    copyFolderEmailIcon.innerHTML = originalIcon;
                    copyFolderEmailIcon.title = originalTitle;
                }, 1500);
            } catch (err) {
                console.error('Failed to copy HTML to clipboard: ', err);
                try {
                    await navigator.clipboard.writeText(plainTextToCopy.trim());
                    const originalIcon = copyFolderEmailIcon.innerHTML;
                    const originalTitle = copyFolderEmailIcon.title;
                    copyFolderEmailIcon.innerHTML = '⚠️';
                    copyFolderEmailIcon.title = 'Copied as plain text!';
                    setTimeout(() => {
                        copyFolderEmailIcon.innerHTML = originalIcon;
                        copyFolderEmailIcon.title = originalTitle;
                    }, 2000);
                } catch (textErr) {
                    console.error('Failed to copy plain text to clipboard: ', textErr);
                    alert('Failed to copy folder details. See console for errors.');
                }
            }
        };
        actionsContainer.appendChild(copyFolderEmailIcon);
        header.appendChild(actionsContainer);
    }
    li.appendChild(header);
    if (isExpanded) {
        const childrenUl = document.createElement('ul');
        childrenUl.classList.add('ml-5', 'pl-3', 'border-l-2', 'border-gray-200');
        folders.filter(f => f.parentId === folder.id && !f.isDeleted).sort((a, b) => a.name.localeCompare(b.name)).forEach(cf => childrenUl.appendChild(createFolderElement(cf)));
        files.filter(f => f.folderId === folder.id && !f.isDeleted).sort((a, b) => a.currentName.localeCompare(b.currentName)).forEach(cf => childrenUl.appendChild(createFileElement(cf)));
        if (childrenUl.children.length > 0) {
            li.appendChild(childrenUl);
        } else {
            const emptyMsg = document.createElement('li');
            emptyMsg.textContent = "Empty folder";
            emptyMsg.classList.add('text-xs', 'text-muted-text', 'italic', 'py-1');
            const emptyMsgListContainer = document.createElement('ul');
            emptyMsgListContainer.classList.add('ml-5', 'pl-3');
            emptyMsgListContainer.appendChild(emptyMsg);
            li.appendChild(emptyMsgListContainer);
        }
    }
    return li;
}

function toggleFolder(folderId, ensureExpanded = null) {
    const index = expandedFolders.indexOf(folderId);
    let changed = false;
    if (ensureExpanded === true) {
        if (index === -1) { expandedFolders.push(folderId); changed = true; }
    } else if (ensureExpanded === false) {
        if (index > -1) { expandedFolders.splice(index, 1); changed = true; }
    } else {
        if (index > -1) expandedFolders.splice(index, 1);
        else expandedFolders.push(folderId);
        changed = true;
    }

    if (changed) {
        _commitLocalStateChanges();
    }
    renderFileTree();
}

// --- MODIFIED createFileElement ---
function createFileElement(file) {
    if (file.isDeleted) return document.createDocumentFragment();
    const li = document.createElement('li');
    li.classList.add('file', 'py-0.5', 'file-tree-item-hover');
    li.dataset.fileId = file.id; // file.id is SHA256 here

    const wrapper = document.createElement('div');
    wrapper.classList.add('file-item-wrapper', 'flex', 'items-center', 'p-1.5', 'rounded-md');

    const iconSpan = document.createElement('span');
    iconSpan.textContent = '📄';
    iconSpan.classList.add('mr-2', 'text-gray-500', 'text-lg');
    wrapper.appendChild(iconSpan);

    const nameSpan = document.createElement('span');
    nameSpan.classList.add('text-sm', 'text-gray-600', 'truncate', 'flex-grow', 'cursor-pointer');
    nameSpan.addEventListener('click', (e) => { e.stopPropagation(); openMetadataModal(file.id); });

    const activeSharesForFileCount = shares.filter(s => {
        if (s.itemId === file.id && s.itemType === 'file') {
            const sharer = sharers.find(sh => sh.id === s.sharerId);
            if (sharer && !sharer.stopAllShares) {
                if (!s.expiryDate || new Date(s.expiryDate) > new Date()) { return true; }
            }
        }
        return false;
    }).length;

    nameSpan.textContent = file.currentName;
    if (activeSharesForFileCount > 0) {
        nameSpan.textContent += ` (${activeSharesForFileCount})`;
        nameSpan.title = `${file.currentName} - Shared with ${activeSharesForFileCount} user(s)`;
    } else {
        nameSpan.title = file.currentName;
    }
    if (file.isPublic && file.driveFileId) {
        nameSpan.innerHTML += ' <span class="text-xs text-brand-green" title="Publicly viewable via Drive link">🌍</span>';
    }


    if (file.encryptedDriveJsonId) {
        const encryptedIconSpan = document.createElement('span');
        encryptedIconSpan.classList.add('inline-block', 'ml-1', 'cursor-pointer', 'hover:opacity-75');
        encryptedIconSpan.title = `Cache & Open Encrypted Preview for ${file.currentName}`;
        encryptedIconSpan.innerHTML = '🔒';
        encryptedIconSpan.onclick = async (e) => {
            e.preventDefault();
            e.stopPropagation();

            if (encryptedIconSpan.innerHTML.includes('🔄') || encryptedIconSpan.innerHTML === '🔄') {
                console.log("Action already in progress (tree view).");
                return;
            }
            encryptedIconSpan.innerHTML = '🔄';

            console.log(`[UI_CLICK_TREE] Clicked padlock for file: ${file.currentName}, GDriveEncID: ${file.encryptedDriveJsonId}, LocalID (SHA256): ${file.id}, ODFragID: ${file.oneDriveFragmentId}, ODPwdID: ${file.oneDrivePasswordItemId}`);
            await fetchAndCacheEncryptedJsonForPreview(file.encryptedDriveJsonId, file.id, file.currentName);
            encryptedIconSpan.innerHTML = '🔒';
        };
        nameSpan.appendChild(encryptedIconSpan);
    }

    wrapper.appendChild(nameSpan);

    // --- START: ADDED ACTION ICONS ---
    const actionsContainer = document.createElement('div');
    actionsContainer.classList.add('ml-auto', 'flex', 'items-center', 'space-x-1', 'pl-1');

    const viewIcon = document.createElement('span');
    viewIcon.innerHTML = '👁️';
    viewIcon.classList.add('cursor-pointer', 'action-icon', 'text-xs', 'hover:text-brand-blue', 'px-1');
    viewIcon.title = "View file in Google Drive";
    viewIcon.onclick = (e) => { e.stopPropagation(); openViewFileModal(file.id); };
    actionsContainer.appendChild(viewIcon);

    const editIcon = document.createElement('span');
    editIcon.textContent = '⚙️';
    editIcon.classList.add('cursor-pointer', 'action-icon', 'text-xs', 'hover:text-brand-blue', 'px-1');
    editIcon.title = "Edit metadata / Share";
    editIcon.onclick = (e) => { e.stopPropagation(); openMetadataModal(file.id); };
    actionsContainer.appendChild(editIcon);

    const copyLinkIcon = document.createElement('span');
    copyLinkIcon.innerHTML = '🔗';
    copyLinkIcon.classList.add('cursor-pointer', 'action-icon', 'text-xs', 'hover:text-brand-blue', 'px-1');
    copyLinkIcon.title = "Copy shareable link (app internal)";
    copyLinkIcon.onclick = (e) => {
        e.stopPropagation();
        const urlToCopy = `${window.location.origin}${window.location.pathname}?fileid=${encodeURIComponent(file.id)}`;
        navigator.clipboard.writeText(urlToCopy).then(() => {
            const originalIcon = copyLinkIcon.innerHTML;
            const originalTitle = copyLinkIcon.title;
            copyLinkIcon.innerHTML = '✅';
            copyLinkIcon.title = 'Link Copied!';
            setTimeout(() => {
                copyLinkIcon.innerHTML = originalIcon;
                copyLinkIcon.title = originalTitle;
            }, 1500);
        }).catch(err => {
            console.error('Failed to copy link: ', err);
            alert('Failed to copy link. See console for details.');
        });
    };
    actionsContainer.appendChild(copyLinkIcon);

    const copyEmailIcon = document.createElement('span');
    copyEmailIcon.innerHTML = '📧';
    copyEmailIcon.classList.add('cursor-pointer', 'action-icon', 'text-xs', 'hover:text-brand-blue', 'px-1');
    copyEmailIcon.title = "Copy file details for email";
    copyEmailIcon.onclick = async (e) => {
        e.stopPropagation();
        const currentFile = files.find(f => f.id === file.id && !f.isDeleted);
        if (!currentFile) { alert("File not found."); return; }

        const appUrl = `${window.location.origin}${window.location.pathname}?fileid=${encodeURIComponent(currentFile.id)}`;
        const driveUrl = currentFile.driveFileId ? `https://drive.google.com/file/d/${currentFile.driveFileId}/view?usp=sharing` : "N/A (Not synced to Drive)";

        let sharedWithDetailsText = [];
        let sharedWithDetailsHtmlList = [];
        const activeSharesForThisFile = shares.filter(s => {
            if (s.itemId === currentFile.id && s.itemType === 'file') {
                const sharer = sharers.find(sh => sh.id === s.sharerId);
                if (sharer && !sharer.stopAllShares) {
                    if (!s.expiryDate || new Date(s.expiryDate) > new Date()) { return true; }
                }
            }
            return false;
        });
        activeSharesForThisFile.forEach(share => {
            const sharer = sharers.find(sh => sh.id === share.sharerId);
            if (sharer) {
                const detailString = `${sharer.fullname} (${sharer.email})`;
                sharedWithDetailsText.push(detailString);
                sharedWithDetailsHtmlList.push(`<li>${detailString.replace(/</g, "<").replace(/>/g, ">")}</li>`);
            }
        });

        let sharedWithHtmlContent = sharedWithDetailsHtmlList.length > 0 ? `<ul>${sharedWithDetailsHtmlList.join('')}</ul>` : '<p style="font-style: italic;">Not currently shared with anyone directly via this system.</p>';
        let sharedWithPlainTextContent = sharedWithDetailsText.length > 0 ? sharedWithDetailsText.map(s => `- ${s}`).join('\n') : 'Not currently shared with anyone directly via this system.';
        let publicAccessText = currentFile.isPublic && currentFile.driveFileId ? 'Publicly Viewable (anyone with Drive link)' : 'Private';
        let publicAccessHtml = currentFile.isPublic && currentFile.driveFileId ? '<span style="color: green; font-weight: bold;">Publicly Viewable</span> (anyone with Drive link)' : 'Private';

        const htmlToCopy = `<div style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 1.6;"><p>Here are the details for the file:</p><p><strong>File Name:</strong> ${currentFile.currentName.replace(/</g, "<").replace(/>/g, ">")}<br><strong>App Link:</strong> <a href="${appUrl}">${appUrl}</a><br><strong>Google Drive Link:</strong> <a href="${driveUrl}">${driveUrl}</a><br><strong>Public Access:</strong> ${publicAccessHtml}</p><p><strong>Currently Shared With (via app):</strong></p>${sharedWithHtmlContent}</div>`;
        const plainTextToCopy = `File details:\nFile Name: ${currentFile.currentName}\nApp Link: ${appUrl}\nGoogle Drive Link: ${driveUrl}\nPublic Access: ${publicAccessText}\n\nCurrently Shared With (via app):\n${sharedWithPlainTextContent}`;

        try {
            const htmlBlob = new Blob([htmlToCopy], { type: 'text/html' });
            const textBlob = new Blob([plainTextToCopy.trim()], { type: 'text/plain' });
            const clipboardItem = new ClipboardItem({ 'text/html': htmlBlob, 'text/plain': textBlob });
            await navigator.clipboard.write([clipboardItem]);
            const originalIcon = copyEmailIcon.innerHTML;
            const originalTitle = copyEmailIcon.title;
            copyEmailIcon.innerHTML = '✅';
            copyEmailIcon.title = 'Details Copied!';
            setTimeout(() => {
                copyEmailIcon.innerHTML = originalIcon;
                copyEmailIcon.title = originalTitle;
            }, 1500);
        } catch (err) {
            console.error('Failed to copy HTML to clipboard: ', err);
            try {
                await navigator.clipboard.writeText(plainTextToCopy.trim());
                const originalIcon = copyEmailIcon.innerHTML;
                const originalTitle = copyEmailIcon.title;
                copyEmailIcon.innerHTML = '⚠️';
                copyEmailIcon.title = 'Copied as plain text!';
                setTimeout(() => {
                    copyEmailIcon.innerHTML = originalIcon;
                    copyEmailIcon.title = originalTitle;
                }, 2000);
            } catch (textErr) {
                console.error('Failed to copy plain text to clipboard: ', textErr);
                alert('Failed to copy file details. See console for errors.');
            }
        }
    };
    actionsContainer.appendChild(copyEmailIcon);
    wrapper.appendChild(actionsContainer);
    // --- END: ADDED ACTION ICONS ---

    li.appendChild(wrapper);
    wrapper.setAttribute('draggable', 'true');
    wrapper.addEventListener('dragstart', (e) => handleFileDragStart(e, file.id));
    wrapper.addEventListener('dragend', handleDragEndCommon);
    return li;
}

// --- MODIFIED renderDataTable ---
function renderDataTable(searchTerm = '') {
    dataTableBody.innerHTML = '';
    const lowerSearchTerm = searchTerm.toLowerCase();
    let displayFiles = files.filter(file => !file.isDeleted);

    if (searchTerm) {
        displayFiles = displayFiles.filter(file => {
            const folder = folders.find(f => f.id === file.folderId && !f.isDeleted);
            const folderName = folder ? folder.name : (file.folderId === 'root' ? 'Root' : 'N/A');
            return file.currentName.toLowerCase().includes(lowerSearchTerm) ||
                file.originalName.toLowerCase().includes(lowerSearchTerm) ||
                folderName.toLowerCase().includes(lowerSearchTerm) ||
                (file.tags && file.tags.join(',').toLowerCase().includes(lowerSearchTerm)) ||
                (file.comments && file.comments.toLowerCase().includes(lowerSearchTerm)) ||
                file.mimeType.toLowerCase().includes(lowerSearchTerm);
        });
    }

    if (filesSortColumn) {
        displayFiles.sort((a, b) => {
            let valA, valB;
            if (filesSortColumn === 'folderName') {
                valA = (folders.find(f => f.id === a.folderId && !f.isDeleted)?.name || '').toLowerCase();
                valB = (folders.find(f => f.id === b.folderId && !f.isDeleted)?.name || '').toLowerCase();
            } else {
                valA = a[filesSortColumn];
                valB = b[filesSortColumn];
                if (typeof valA === 'string') valA = valA.toLowerCase();
                if (typeof valB === 'string') valB = valB.toLowerCase();
            }
            if (filesSortColumn === 'size') { valA = Number(valA); valB = Number(valB); }
            if (filesSortColumn === 'uploadDate') { valA = new Date(valA); valB = new Date(valB); }
            if (valA < valB) return filesSortDirection === 'asc' ? -1 : 1;
            if (valA > valB) return filesSortDirection === 'asc' ? 1 : -1;
            return 0;
        });
    }

    if (filesDataTableHeaders) {
        filesDataTableHeaders.forEach(th => {
            th.classList.remove('sort-asc', 'sort-desc');
            if (th.dataset.sort === filesSortColumn) {
                th.classList.add(filesSortDirection === 'asc' ? 'sort-asc' : 'sort-desc');
            }
        });
    }


    const totalItems = displayFiles.length;
    const totalPages = Math.max(1, Math.ceil(totalItems / itemsPerPageFiles));
    filesDataTablePage = Math.max(1, Math.min(filesDataTablePage, totalPages));
    const startIndex = (filesDataTablePage - 1) * itemsPerPageFiles;
    const endIndex = startIndex + itemsPerPageFiles;
    const paginatedFiles = displayFiles.slice(startIndex, endIndex);

    if (totalItems === 0) {
        const tr = dataTableBody.insertRow();
        const td = tr.insertCell();
        td.colSpan = 7;
        td.textContent = searchTerm ? 'No files match search.' : 'No files available.';
        td.classList.add('text-center', 'text-muted-text', 'py-4');
    } else if (paginatedFiles.length === 0 && totalItems > 0) {
        const tr = dataTableBody.insertRow();
        const td = tr.insertCell();
        td.colSpan = 7;
        td.textContent = 'No files on this page.';
        td.classList.add('text-center', 'text-muted-text', 'py-4');
    } else {
        paginatedFiles.forEach(file => {
            const tr = dataTableBody.insertRow();
            tr.dataset.fileId = file.id;
            const folderObject = folders.find(f => f.id === file.folderId && !f.isDeleted);
            const folderName = folderObject ? folderObject.name : (file.folderId === 'root' ? 'Root' : 'Unknown');

            const nameCell = tr.insertCell();
            nameCell.textContent = file.currentName;
            if (file.isPublic && file.driveFileId) {
                const publicIcon = document.createElement('span');
                publicIcon.textContent = ' 🌍';
                publicIcon.title = 'Publicly viewable via Drive link';
                publicIcon.classList.add('text-xs', 'text-brand-green');
                nameCell.appendChild(publicIcon);
            }

            if (file.encryptedDriveJsonId) {
                const encryptedIconSpan = document.createElement('span');
                encryptedIconSpan.classList.add('text-xs', 'text-brand-blue', 'cursor-pointer', 'hover:opacity-75', 'ml-1');
                encryptedIconSpan.title = `Cache Encrypted Data for Preview (GDrive ID: ${file.encryptedDriveJsonId}${file.oneDriveFragmentId ? ', OD Frag ID: ' + file.oneDriveFragmentId : ''}${file.oneDrivePasswordItemId ? ', OD Pwd ID: ' + file.oneDrivePasswordItemId : ''})`;
                encryptedIconSpan.innerHTML = '🔒';
                encryptedIconSpan.onclick = async (e) => {
                    e.preventDefault();
                    e.stopPropagation();

                    if (encryptedIconSpan.innerHTML.includes('🔄') || encryptedIconSpan.innerHTML === '🔄') {
                        console.log("Action already in progress (data table).");
                        return;
                    }
                    encryptedIconSpan.innerHTML = '🔄';
                    console.log(`[UI_CLICK_TABLE] Clicked cache icon for file: ${file.currentName}, GDriveEncID: ${file.encryptedDriveJsonId}, LocalID: ${file.id}, ODFragID: ${file.oneDriveFragmentId}, ODPwdID: ${file.oneDrivePasswordItemId}`);
                    await fetchAndCacheEncryptedJsonForPreview(file.encryptedDriveJsonId, file.id, file.currentName);
                    encryptedIconSpan.innerHTML = '🔒';
                };
                nameCell.appendChild(encryptedIconSpan);
            }

            tr.insertCell().textContent = folderName;
            tr.insertCell().textContent = (file.tags && file.tags.join(', ')) || '-';
            tr.insertCell().textContent = formatBytes(file.size);
            tr.insertCell().textContent = file.mimeType;
            tr.insertCell().textContent = new Date(file.uploadDate).toLocaleDateString();

            const actionsTd = tr.insertCell();
            actionsTd.classList.add('whitespace-nowrap');
            const editButton = document.createElement('button');
            editButton.innerHTML = '✏️ <span class="sr-only">Edit</span>';
            editButton.classList.add('text-brand-blue', 'hover:text-blue-700', 'p-1', 'text-lg');
            editButton.title = "Edit metadata / Share";
            editButton.onclick = () => openMetadataModal(file.id);
            actionsTd.appendChild(editButton);
        });
    }
    updatePaginationControls(paginationControlsFiles, pageInfoFiles, prevPageBtnFiles, nextPageBtnFiles, totalItems, totalPages, filesDataTablePage);
}


function updatePaginationControls(controlsContainer, pageInfoEl, prevBtn, nextBtn, totalItems, totalPages, currentPage) {
    if (totalItems === 0 && currentPage === 1 && totalPages === 1) {
        controlsContainer.classList.add('hidden');
        const itemsPerPageSelector = controlsContainer.parentElement.querySelector('select[id^="itemsPerPage"]');
        if (itemsPerPageSelector && itemsPerPageSelector.parentElement) itemsPerPageSelector.parentElement.classList.add('hidden');
        return;
    }
    controlsContainer.classList.remove('hidden');
    const itemsPerPageSelector = controlsContainer.parentElement.querySelector('select[id^="itemsPerPage"]');
    if (itemsPerPageSelector && itemsPerPageSelector.parentElement) itemsPerPageSelector.parentElement.classList.remove('hidden');

    pageInfoEl.textContent = `Page ${currentPage} of ${totalPages}`;
    prevBtn.disabled = currentPage === 1;
    nextBtn.disabled = currentPage === totalPages;
}

function getFolderDisplayPathForDropdown(folderId, currentFoldersRef = folders) { const path = []; let currentId = folderId; let safety = 0; const maxDepth = currentFoldersRef.length + 5; while (currentId && safety < maxDepth) { const folder = currentFoldersRef.find(f => f.id === currentId && !f.isDeleted); if (folder) { path.unshift(folder.name); currentId = folder.parentId; } else { if (currentId === 'root' && !path.find(pName => pName.toLowerCase() === 'root')) path.unshift('Root'); break; } safety++; } if (path.length === 0) { const folderCheck = currentFoldersRef.find(f => f.id === folderId && !f.isDeleted); if (folderId === 'root' || !folderCheck) return 'Root'; } return path.join(' / '); }
function populateParentFolderDropdown() { modalParentFolder.innerHTML = ''; const allValidFolders = folders.filter(f => !f.isDeleted); allValidFolders.sort((a, b) => getFolderDisplayPathForDropdown(a.id, allValidFolders).localeCompare(getFolderDisplayPathForDropdown(b.id, allValidFolders))); allValidFolders.forEach(folder => { const option = document.createElement('option'); option.value = folder.id; option.textContent = getFolderDisplayPathForDropdown(folder.id, allValidFolders); modalParentFolder.appendChild(option); }); const activeFolderExists = folders.find(f => f.id === activeFolderId && !f.isDeleted); if (activeFolderExists) modalParentFolder.value = activeFolderId; else if (allValidFolders.length > 0) modalParentFolder.value = allValidFolders[0].id; }
function openNewFolderModal() { populateParentFolderDropdown(); modalNewFolderName.value = ''; newFolderModal.classList.remove('hidden'); modalNewFolderName.focus(); }

modalNewFolderCreateBtn.addEventListener('click', () => {
    const folderName = modalNewFolderName.value.trim();
    const parentId = modalParentFolder.value;
    if (!folderName) { alert('Folder name cannot be empty.'); modalNewFolderName.focus(); return; }
    if (/[\\/:*?"<>|]/.test(folderName)) { alert('Folder name contains invalid characters.'); modalNewFolderName.focus(); return; }
    if (folders.some(f => f.parentId === parentId && !f.isDeleted && f.name.toLowerCase() === folderName.toLowerCase())) { alert(`A folder named "${folderName}" already exists in this location.`); modalNewFolderName.focus(); return; }

    const newFolder = { id: 'folder_' + generateUUID(), name: folderName, parentId: parentId, isDeleted: false, deletedDate: null };
    folders.push(newFolder);
    if (!expandedFolders.includes(parentId)) expandedFolders.push(parentId);
    activeFolderId = newFolder.id;
    ensurePathExpanded(activeFolderId);

    _commitLocalStateChanges();
    renderAll();
    newFolderModal.classList.add('hidden');
});

function openRenameFolderModal(folderIdToManage) { const folder = folders.find(f => f.id === folderIdToManage && !f.isDeleted); if (!folder || folder.id === 'root') { alert("Cannot manage or share the Root folder directly this way."); return; } modalRenameFolderId.value = folder.id; modalRenameFolderNameInput.value = folder.name; modalRenameFolderParentInfo.textContent = getFolderDisplayPathForDropdown(folder.parentId) || 'N/A'; renderFolderShareControls(folder.id); renameFolderModal.classList.remove('hidden'); modalRenameFolderNameInput.focus(); }

modalRenameFolderSaveBtn.addEventListener('click', () => {
    const folderId = modalRenameFolderId.value;
    const newName = modalRenameFolderNameInput.value.trim();
    const folder = folders.find(f => f.id === folderId && !f.isDeleted);
    if (!folder) { alert("Folder not found. Cannot save changes."); renameFolderModal.classList.add('hidden'); return; }
    if (!newName) { alert("Folder name cannot be empty."); modalRenameFolderNameInput.focus(); return; }
    if (/[\\/:*?"<>|]/.test(newName)) { alert('Folder name contains invalid characters.'); modalRenameFolderNameInput.focus(); return; }
    if (folders.some(f => f.parentId === folder.parentId && f.id !== folderId && f.name.toLowerCase() === newName.toLowerCase() && !f.isDeleted)) { alert(`A folder named "${newName}" already exists in this location.`); modalRenameFolderNameInput.focus(); return; }

    if (folder.name !== newName) folder.name = newName;
    saveFolderShares(folder.id);

    _commitLocalStateChanges();
    renderAll();
    renameFolderModal.classList.add('hidden');
});

modalDeleteFolderBtn.addEventListener('click', async () => {
    const folderIdToDelete = modalRenameFolderId.value;
    const folderToDelete = folders.find(f => f.id === folderIdToDelete);
    if (!folderToDelete || folderToDelete.id === 'root' || folderToDelete.isDeleted) { alert("Cannot delete this folder or it's already deleted."); renameFolderModal.classList.add('hidden'); return; }

    if (confirm(`Are you sure you want to delete the folder "${folderToDelete.name}" and all its contents? This will also remove its shares and attempt to delete associated files from Google Drive and OneDrive. This action cannot be undone.`)) {
        const parentOfDeleted = folderToDelete.parentId;

        await flagFolderAndContentsAsDeleted(folderIdToDelete);

        let allDeletedItemIds = new Set();
        function collectDeletedIdsRecursive(fId) {
            allDeletedItemIds.add(fId);
            folders.filter(sf => sf.parentId === fId).forEach(subF => collectDeletedIdsRecursive(subF.id));
            files.filter(fi => fi.folderId === fId).forEach(itemFile => allDeletedItemIds.add(itemFile.id));
        }
        collectDeletedIdsRecursive(folderIdToDelete);
        shares = shares.filter(s => !allDeletedItemIds.has(s.itemId));

        _commitLocalStateChanges();

        const currentActiveStillValid = folders.find(f => f.id === activeFolderId && !f.isDeleted);
        if (!currentActiveStillValid) {
            activeFolderId = folders.find(f => f.id === parentOfDeleted && !f.isDeleted) ? parentOfDeleted : 'root';
            _commitLocalStateChanges();
        }
        renderAll();
        renameFolderModal.classList.add('hidden');
    }
});

// MODIFIED: Now cleans up cached data from IndexedDB
async function flagFolderAndContentsAsDeleted(folderId) {
    if (folderId === 'root') return;
    const now = new Date().toISOString();

    const folder = folders.find(f => f.id === folderId);
    if (folder && !folder.isDeleted) {
        folder.isDeleted = true;
        folder.deletedDate = now;
    }

    const filesToDeletePromises = [];
    files.filter(file => file.folderId === folderId && !file.isDeleted)
        .forEach(file => {
            file.isDeleted = true;
            file.deletedDate = now;

            // ADDED: Clean up cached data from IndexedDB during deletion
            const gDriveKey = file.encryptedDriveJsonId ? `encryptedJsonContent_${file.encryptedDriveJsonId}` : null;
            const odFragKey = file.oneDriveFragmentId ? `oneDriveFragmentContent_${file.oneDriveFragmentId}` : null;
            const odPwdKey = file.oneDrivePasswordItemId ? `passwordContent_${file.oneDrivePasswordItemId}` : null;
            const fragInfoKey = `fragInfo_${file.id}`; // file.id is sha256
            if (gDriveKey) dbHelper.delete(gDriveKey).catch(e => console.warn(`DB cleanup failed for ${gDriveKey}`, e));
            if (odFragKey) dbHelper.delete(odFragKey).catch(e => console.warn(`DB cleanup failed for ${odFragKey}`, e));
            if (odPwdKey) dbHelper.delete(odPwdKey).catch(e => console.warn(`DB cleanup failed for ${odPwdKey}`, e));
            dbHelper.delete(fragInfoKey).catch(e => console.warn(`DB cleanup failed for ${fragInfoKey}`, e));


            if (file.driveFileId) {
                filesToDeletePromises.push(deleteDriveFileById(file.driveFileId).catch(e => { console.warn(`Cloud delete failed for GDrive file ${file.driveFileId}: ${e.message}`); return { status: 'rejected' }; }));
            }
            if (file.encryptedDriveJsonId) {
                filesToDeletePromises.push(deleteDriveFileById(file.encryptedDriveJsonId).catch(e => { console.warn(`Cloud delete failed for GDrive JSON ${file.encryptedDriveJsonId}: ${e.message}`); return { status: 'rejected' }; }));
            }
            if (file.oneDriveFragmentId) {
                filesToDeletePromises.push(deleteOneDriveFileByItemId(file.oneDriveFragmentId).catch(e => { console.warn(`Cloud delete failed for OneDrive Frag ${file.oneDriveFragmentId}: ${e.message}`); return { status: 'rejected' }; }));
            }
            if (file.oneDrivePasswordItemId) {
                filesToDeletePromises.push(deleteOneDriveFileByItemId(file.oneDrivePasswordItemId).catch(e => { console.warn(`Cloud delete failed for OneDrive Pwd ${file.oneDrivePasswordItemId}: ${e.message}`); return { status: 'rejected' }; }));
            }
        });
    await Promise.allSettled(filesToDeletePromises);

    const subfoldersToDelete = folders.filter(f => f.parentId === folderId && !f.isDeleted);
    for (const subfolder of subfoldersToDelete) {
        await flagFolderAndContentsAsDeleted(subfolder.id);
    }
}


function handleDragEndCommon(event) { draggedItemId = null; draggedItemType = null; document.querySelectorAll('.drop-target-highlight').forEach(el => el.classList.remove('drop-target-highlight')); if (event.target.getAttribute('draggable') === 'true') event.target.style.cursor = 'grab'; }
function handleFileDragStart(event, fileId) { const file = files.find(f => f.id === fileId && !f.isDeleted); if (!file) { event.preventDefault(); return; } draggedItemId = fileId; draggedItemType = 'file'; event.dataTransfer.setData('text/plain', fileId); event.dataTransfer.effectAllowed = 'move'; event.target.style.cursor = 'grabbing'; }
function handleFolderDragStart(event, folderId) { const folder = folders.find(f => f.id === folderId && !f.isDeleted); if (!folder || folder.id === 'root') { event.preventDefault(); return; } draggedItemId = folderId; draggedItemType = 'folder'; event.dataTransfer.setData('text/plain', folderId); event.dataTransfer.effectAllowed = 'move'; event.target.style.cursor = 'grabbing'; }
function isDescendant(potentialChildId, potentialParentId) { if (potentialChildId === potentialParentId) return true; let currentFolder = folders.find(f => f.id === potentialChildId && !f.isDeleted); let safety = 0; while (currentFolder && currentFolder.parentId && safety < folders.length + 5) { if (currentFolder.parentId === potentialParentId) return true; currentFolder = folders.find(f => f.id === currentFolder.parentId && !f.isDeleted); safety++; } return false; }

function handleDropOnFolder(event, targetFolderId) {
    event.preventDefault();
    event.stopPropagation();
    document.querySelectorAll('.drop-target-highlight').forEach(el => el.classList.remove('drop-target-highlight'));
    const targetFolder = folders.find(f => f.id === targetFolderId && !f.isDeleted);
    if (!targetFolder || !draggedItemId || !draggedItemType) {
        draggedItemId = null; draggedItemType = null; return;
    }

    let changed = false;
    if (draggedItemType === 'file') {
        const fileToMove = files.find(f => f.id === draggedItemId && !f.isDeleted);
        if (fileToMove && fileToMove.folderId !== targetFolderId) {
            if (files.some(f => f.folderId === targetFolderId && f.currentName.toLowerCase() === fileToMove.currentName.toLowerCase() && !f.isDeleted)) {
                alert(`A file named "${fileToMove.currentName}" already exists in the target folder.`);
            } else {
                fileToMove.folderId = targetFolderId;
                changed = true;
            }
        }
    } else if (draggedItemType === 'folder') {
        const folderToMove = folders.find(f => f.id === draggedItemId && !f.isDeleted);
        if (folderToMove && folderToMove.id !== targetFolderId && folderToMove.parentId !== targetFolderId && !isDescendant(targetFolderId, folderToMove.id)) {
            if (folders.some(f => f.parentId === targetFolderId && f.name.toLowerCase() === folderToMove.name.toLowerCase() && f.id !== folderToMove.id && !f.isDeleted)) {
                alert(`A folder named "${folderToMove.name}" already exists in the target location.`);
            } else {
                folderToMove.parentId = targetFolderId;
                changed = true;
            }
        } else if (isDescendant(targetFolderId, folderToMove.id)) {
            alert("Cannot move a folder into one of its own subfolders.");
        }
    }

    draggedItemId = null;
    draggedItemType = null;

    if (changed) {
        if (ensurePathExpanded(targetFolderId)) changed = true;
        _commitLocalStateChanges();
        renderAll();
    }
}

function openMetadataModal(fileIdToEdit) {
    const file = files.find(f => f.id === fileIdToEdit && !f.isDeleted);
    if (!file) { alert("File not found."); return; }
    modalFileId.value = file.id;
    modalFileName.value = file.currentName;
    modalFileTags.value = (file.tags && file.tags.join(', ')) || '';
    modalFileComments.value = file.comments || '';

    const publicAccessContainer = modalFilePublicAccess.closest('.border-t');

    if (file.driveFileId && isDriveAuthenticated) {
        publicAccessContainer.style.display = 'block';
        modalFilePublicAccess.checked = !!file.isPublic;
        modalFilePublicAccess.disabled = !isDriveReadyForOps;

        if (file.isPublic) {
            modalPublicLinkStatus.innerHTML = `Currently: <span class="font-semibold text-brand-green">Public</span>.
                    <a href="https://drive.google.com/file/d/${file.driveFileId}/view?usp=sharing" target="_blank" class="text-brand-blue hover:underline">View Link</a>`;
        } else {
            modalPublicLinkStatus.textContent = 'Currently: Private.';
        }
        if (!isDriveReadyForOps) {
            modalPublicLinkStatus.innerHTML += ' <span class="text-brand-red">(Drive operations pending, toggle disabled)</span>';
        }
    } else {
        publicAccessContainer.style.display = 'none';
        modalFilePublicAccess.checked = false;
        modalPublicLinkStatus.textContent = '';
    }

    renderFileShareControls(file.id);
    metadataModal.classList.remove('hidden');
    modalFileName.focus();
}

modalSaveBtn.addEventListener('click', async () => {
    const fileId = modalFileId.value;
    const file = files.find(f => f.id === fileId && !f.isDeleted);
    if (!file) {
        alert("File not found. Cannot save changes.");
        metadataModal.classList.add('hidden');
        return;
    }

    const newName = modalFileName.value.trim();
    if (!newName) {
        alert("File name cannot be empty.");
        modalFileName.focus();
        return;
    }
    if (/[\\/:*?"<>|]/.test(newName)) {
        alert('File name contains invalid characters.');
        modalFileName.focus();
        return;
    }
    if (newName !== file.currentName && files.some(f => f.folderId === file.folderId && f.id !== file.id && f.currentName.toLowerCase() === newName.toLowerCase() && !f.isDeleted)) {
        alert(`A file named "${newName}" already exists in this folder.`);
        modalFileName.focus();
        return;
    }

    const originalIsPublic = !!file.isPublic;
    let publicStatusSuccessfullyChangedOnDrive = true;

    if (file.driveFileId && isDriveReadyForOps) {
        const wantsPublic = modalFilePublicAccess.checked;
        if (wantsPublic !== originalIsPublic) {
            try {
                const originalStatusText = googleApiStatusEl.textContent;
                googleApiStatusEl.textContent = "Updating public access settings on Drive...";
                if (wantsPublic) {
                    await setFilePublicOnDrive(file.driveFileId);
                    file.isPublic = true;
                    googleApiStatusEl.textContent = "File made public on Drive.";
                } else {
                    await setFilePrivateOnDrive(file.driveFileId);
                    file.isPublic = false;
                    googleApiStatusEl.textContent = "File made private on Drive.";
                }
                setTimeout(() => {
                    if (googleApiStatusEl.textContent.startsWith("File made")) {
                        googleApiStatusEl.textContent = originalStatusText;
                    }
                }, 5000);

            } catch (error) {
                publicStatusSuccessfullyChangedOnDrive = false;
                file.isPublic = originalIsPublic;
                modalFilePublicAccess.checked = originalIsPublic;

                console.error("Error updating public access on Drive:", error);
                alert(`Failed to update public access on Google Drive: ${error.message}.\nThe public status change has been reverted.`);
                googleApiStatusEl.textContent = `Error updating public access: ${error.message}`;
            }
        }
    } else if (modalFilePublicAccess.closest('.border-t').style.display !== 'none' && modalFilePublicAccess.checked !== originalIsPublic) {
        alert("Cannot change public access status. File not synced with Google Drive or Drive is not ready. Reverting toggle.");
        modalFilePublicAccess.checked = originalIsPublic;
        publicStatusSuccessfullyChangedOnDrive = false;
    }

    file.currentName = newName;
    file.tags = modalFileTags.value.split(',').map(t => t.trim()).filter(t => t);
    file.comments = modalFileComments.value.trim();

    if (!publicStatusSuccessfullyChangedOnDrive) {
        file.isPublic = originalIsPublic;
    }

    saveFileShares(file.id);

    _commitLocalStateChanges();
    renderAll();
    metadataModal.classList.add('hidden');

    if (!googleApiStatusEl.textContent.startsWith("Error updating public access")) {
        if (isDriveReadyForOps && lastSuccessfulSyncTimestamp > 0) {
            setTimeout(() => { googleApiStatusEl.textContent = `Data successfully synced from Google Drive. Last sync: ${new Date(lastSuccessfulSyncTimestamp).toLocaleTimeString()}`; }, 3500);
        } else if (isDriveAuthenticated) {
            setTimeout(() => { googleApiStatusEl.textContent = `Google Drive client initialized. Setting up application folder...`; }, 3500);
        } else {
            setTimeout(() => { googleApiStatusEl.textContent = 'Please sign in to sync with Google Drive. Displaying local data if available.'; }, 3500);
        }
    }
});

modalCancelBtn.addEventListener('click', () => metadataModal.classList.add('hidden'));

// MODIFIED: Now cleans up cached data from IndexedDB on deletion
modalDeleteBtn.addEventListener('click', async () => {
    const fileIdToDelete = modalFileId.value;
    const fileIndex = files.findIndex(f => f.id === fileIdToDelete);
    const file = fileIndex !== -1 ? files[fileIndex] : null;

    if (!file) {
        alert("File not found in local data. It might have been removed already. Refreshing view.");
        metadataModal.classList.add('hidden');
        renderAll();
        return;
    }

    if (file.isDeleted) {
        alert("This file was already marked as deleted. Forcing its removal from local data and refreshing the view.");
        if (fileIndex > -1) {
            files.splice(fileIndex, 1);
        }
        shares = shares.filter(s => !(s.itemId === fileIdToDelete && s.itemType === 'file'));
        _commitLocalStateChanges();
        metadataModal.classList.add('hidden');
        renderAll();
        return;
    }

    if (confirm(`Are you sure you want to delete the file "${file.currentName}"? This will also remove its shares and attempt to delete it from Google Drive and associated OneDrive items (fragment/password).`)) {
        file.isDeleted = true;
        file.deletedDate = new Date().toISOString();

        // ADDED: Clean up cached data from IndexedDB
        const gDriveKey = file.encryptedDriveJsonId ? `encryptedJsonContent_${file.encryptedDriveJsonId}` : null;
        const odFragKey = file.oneDriveFragmentId ? `oneDriveFragmentContent_${file.oneDriveFragmentId}` : null;
        const odPwdKey = file.oneDrivePasswordItemId ? `passwordContent_${file.oneDrivePasswordItemId}` : null;
        const fragInfoKey = `fragInfo_${file.id}`; // file.id is sha256
        const cacheCleanupPromises = [
            gDriveKey ? dbHelper.delete(gDriveKey) : Promise.resolve(),
            odFragKey ? dbHelper.delete(odFragKey) : Promise.resolve(),
            odPwdKey ? dbHelper.delete(odPwdKey) : Promise.resolve(),
            dbHelper.delete(fragInfoKey)
        ].map(p => p.catch(e => console.warn("DB cache cleanup error:", e)));


        // Perform async deletions from cloud services
        const cloudDeletionPromises = [];
        if (file.driveFileId) {
            cloudDeletionPromises.push(deleteDriveFileById(file.driveFileId).catch(e => { console.warn(`Cloud delete failed for GDrive file ${file.driveFileId}: ${e.message}`); }));
        }
        if (file.encryptedDriveJsonId) {
            cloudDeletionPromises.push(deleteDriveFileById(file.encryptedDriveJsonId).catch(e => { console.warn(`Cloud delete failed for GDrive JSON ${file.encryptedDriveJsonId}: ${e.message}`); }));
        }
        if (file.oneDriveFragmentId) {
            cloudDeletionPromises.push(deleteOneDriveFileByItemId(file.oneDriveFragmentId).catch(e => { console.warn(`Cloud delete failed for OneDrive Frag ${file.oneDriveFragmentId}: ${e.message}`); }));
        }
        if (file.oneDrivePasswordItemId) {
            cloudDeletionPromises.push(deleteOneDriveFileByItemId(file.oneDrivePasswordItemId).catch(e => { console.warn(`Cloud delete failed for OneDrive Pwd ${file.oneDrivePasswordItemId}: ${e.message}`); }));
        }

        await Promise.allSettled([...cacheCleanupPromises, ...cloudDeletionPromises]);

        shares = shares.filter(s => !(s.itemId === fileIdToDelete && s.itemType === 'file'));
        _commitLocalStateChanges();
        renderAll();
        metadataModal.classList.add('hidden');
    }
});


modalNewFolderCancelBtn.addEventListener('click', () => newFolderModal.classList.add('hidden'));
modalRenameFolderCancelBtn.addEventListener('click', () => renameFolderModal.classList.add('hidden'));
function formatBytes(bytes, decimals = 2) { if (bytes === undefined || bytes === null || isNaN(parseFloat(bytes)) || !isFinite(bytes)) return 'N/A'; if (bytes === 0) return '0 Bytes'; const k = 1024; const dm = decimals < 0 ? 0 : decimals; const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB']; const i = Math.floor(Math.log(bytes) / Math.log(k)); return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i]; }

// Sharers Tab Functions
function renderSharersTab(searchTerm = '') {
    sharersTableBody.innerHTML = '';
    const lowerSearchTerm = searchTerm.toLowerCase();
    let displaySharers = [...sharers];

    if (searchTerm) {
        displaySharers = displaySharers.filter(sharer =>
            sharer.shortname.toLowerCase().includes(lowerSearchTerm) ||
            sharer.fullname.toLowerCase().includes(lowerSearchTerm) ||
            sharer.email.toLowerCase().includes(lowerSearchTerm) ||
            (sharer.dept && sharer.dept.toLowerCase().includes(lowerSearchTerm))
        );
    }

    if (sharersSortColumn) {
        displaySharers.sort((a, b) => {
            let valA = a[sharersSortColumn];
            let valB = b[sharersSortColumn];
            if (typeof valA === 'string') valA = valA.toLowerCase();
            if (typeof valB === 'string') valB = valB.toLowerCase();
            if (sharersSortColumn === 'dept') { valA = valA || ''; valB = valB || ''; }
            if (valA < valB) return sharersSortDirection === 'asc' ? -1 : 1;
            if (valA > valB) return sharersSortDirection === 'asc' ? 1 : -1;
            return 0;
        });
    }

    if (sharersDataTableHeaders) {
        sharersDataTableHeaders.forEach(th => {
            th.classList.remove('sort-asc', 'sort-desc');
            if (th.dataset.sortSharers === sharersSortColumn) {
                th.classList.add(sharersSortDirection === 'asc' ? 'sort-asc' : 'sort-desc');
            }
        });
    }


    const totalItems = displaySharers.length;
    const totalPages = Math.max(1, Math.ceil(totalItems / itemsPerPageSharers));
    sharersTablePage = Math.max(1, Math.min(sharersTablePage, totalPages));
    const startIndex = (sharersTablePage - 1) * itemsPerPageSharers;
    const endIndex = startIndex + itemsPerPageSharers;
    const paginatedSharers = displaySharers.slice(startIndex, endIndex);

    if (totalItems === 0) {
        const tr = sharersTableBody.insertRow();
        const td = tr.insertCell();
        td.colSpan = 8;
        td.textContent = searchTerm ? 'No sharers match search.' : 'No sharers defined.';
        td.classList.add('text-center', 'text-muted-text', 'py-4');
    } else if (paginatedSharers.length === 0 && totalItems > 0) {
        const tr = sharersTableBody.insertRow();
        const td = tr.insertCell();
        td.colSpan = 8;
        td.textContent = 'No sharers on this page.';
        td.classList.add('text-center', 'text-muted-text', 'py-4');
    } else {
        paginatedSharers.forEach(sharer => {
            const tr = sharersTableBody.insertRow();
            tr.insertCell().textContent = sharer.shortname;
            tr.insertCell().textContent = sharer.fullname;
            tr.insertCell().textContent = sharer.email;
            tr.insertCell().textContent = sharer.dept || '-';
            tr.insertCell().textContent = sharer.type.charAt(0).toUpperCase() + sharer.type.slice(1);
            tr.insertCell().textContent = '••••••••';
            tr.insertCell().textContent = sharer.stopAllShares ? 'Yes' : 'No';

            const actionsTd = tr.insertCell();
            actionsTd.classList.add('whitespace-nowrap');
            const editButton = document.createElement('button');
            editButton.innerHTML = '✏️ <span class="sr-only">Edit</span>';
            editButton.classList.add('text-brand-blue', 'hover:text-blue-700', 'p-1', 'mr-2', 'text-lg');
            editButton.title = "Edit Sharer";
            editButton.onclick = () => openSharerModal(sharer.id);
            actionsTd.appendChild(editButton);

            if (!sharer.isExample) {
                const deleteButton = document.createElement('button');
                deleteButton.innerHTML = '🗑️ <span class="sr-only">Delete</span>';
                deleteButton.classList.add('text-brand-red', 'hover:text-red-700', 'p-1', 'text-lg');
                deleteButton.title = "Delete Sharer";
                deleteButton.onclick = () => {
                    const sharerToDelete = sharers.find(s => s.id === sharer.id);
                    if (sharerToDelete && confirm(`Delete sharer "${sharerToDelete.fullname}" & ALL their shares (files & folders)? This action cannot be undone.`)) {
                        performDeleteSharer(sharer.id);
                    }
                };
                actionsTd.appendChild(deleteButton);
            }
        });
    }
    updatePaginationControls(paginationControlsSharers, pageInfoSharers, prevPageBtnSharers, nextPageBtnSharers, totalItems, totalPages, sharersTablePage);
}
function openSharerModal(sharerIdToEdit = null) { sharerModal.classList.remove('hidden'); if (sharerIdToEdit) { const sharer = sharers.find(s => s.id === sharerIdToEdit); if (!sharer) { alert('Not found!'); sharerModal.classList.add('hidden'); return; } sharerModalTitle.textContent = 'Edit Sharer'; modalSharerId.value = sharer.id; modalSharerShortname.value = sharer.shortname; modalSharerFullname.value = sharer.fullname; modalSharerEmail.value = sharer.email; modalSharerDept.value = sharer.dept || ''; modalSharerType.value = sharer.type; modalSharerPassword.value = ''; modalSharerPassword.placeholder = "Leave blank to keep existing"; modalSharerStopAllShares.checked = sharer.stopAllShares; modalSharerDeleteBtn.style.display = sharer.isExample ? 'none' : 'inline-block'; } else { sharerModalTitle.textContent = 'Add Sharer'; modalSharerId.value = ''; modalSharerShortname.value = ''; modalSharerFullname.value = ''; modalSharerEmail.value = ''; modalSharerDept.value = ''; modalSharerType.value = 'internal'; modalSharerPassword.value = ''; modalSharerPassword.placeholder = "Enter password"; modalSharerStopAllShares.checked = false; modalSharerDeleteBtn.style.display = 'none'; } modalSharerShortname.focus(); }

modalSharerSaveBtn.addEventListener('click', () => {
    const id = modalSharerId.value;
    const shortname = modalSharerShortname.value.trim();
    const fullname = modalSharerFullname.value.trim();
    const email = modalSharerEmail.value.trim().toLowerCase();
    const dept = modalSharerDept.value.trim();
    const type = modalSharerType.value;
    const passwordInput = modalSharerPassword.value;
    const stopAllSharesFlag = modalSharerStopAllShares.checked;
    if (!shortname || !fullname || !email) { alert('Shortname, Full Name, and Email are required.'); return; }
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { alert('Please enter a valid email address.'); modalSharerEmail.focus(); return; }
    if (sharers.some(s => s.id !== id && (s.email === email || s.shortname === shortname))) { alert(`A sharer with this email or shortname already exists.`); return; }

    if (id) {
        const sharer = sharers.find(s => s.id === id);
        if (sharer) {
            sharer.shortname = shortname; sharer.fullname = fullname; sharer.email = email;
            sharer.dept = dept; sharer.type = type;
            if (passwordInput) sharer.password = passwordInput;
            sharer.stopAllShares = stopAllSharesFlag;
        }
    } else {
        if (!passwordInput) { alert('Password is required for new sharers.'); modalSharerPassword.focus(); return; }
        sharers.push({
            id: 'sharer_' + generateUUID(), shortname, fullname, email, dept, type,
            password: passwordInput, stopAllShares: stopAllSharesFlag, isExample: false
        });
    }

    _commitLocalStateChanges();
    renderSharersTab(sharersSearchInput.value);
    renderFileTree();
    renderDataTable(searchInput.value);
    sharerModal.classList.add('hidden');
});

modalSharerCancelBtn.addEventListener('click', () => sharerModal.classList.add('hidden'));

modalSharerDeleteBtn.addEventListener('click', () => {
    const id = modalSharerId.value;
    if (!id) return;
    const sharer = sharers.find(s => s.id === id);
    if (!sharer) { alert("Sharer not found."); return; }
    if (sharer.isExample) { alert("Example sharers cannot be deleted through this modal."); return; }
    if (confirm(`Are you sure you want to delete the sharer "${sharer.fullname}" and ALL their shares (files & folders)? This action cannot be undone.`)) {
        performDeleteSharer(id);
        sharerModal.classList.add('hidden');
    }
});

function performDeleteSharer(sharerId) {
    sharers = sharers.filter(s => s.id !== sharerId);
    shares = shares.filter(sh => sh.sharerId !== sharerId);
    _commitLocalStateChanges();
    renderSharersTab(sharersSearchInput.value);
    renderFileTree();
    renderDataTable(searchInput.value);
}

function renderFileShareControls(fileId) { currentFileModalShares = shares.filter(s => s.itemId === fileId && s.itemType === 'file'); modalFileCurrentlySharedWith.innerHTML = ''; modalFileShareSearch.value = ''; renderFileShareSearchResults(''); modalFileNoSharersMsg.style.display = currentFileModalShares.length === 0 ? 'block' : 'none'; currentFileModalShares.forEach(share => { const sharer = sharers.find(s => s.id === share.sharerId); if (sharer) addFileSharerToSelectedUI(sharer, share.expiryDate); }); }
function renderFileShareSearchResults(searchTerm) { modalFileShareSearchResults.innerHTML = ''; const lowerSearchTerm = searchTerm.toLowerCase(); const alreadySelectedSharerIds = Array.from(modalFileCurrentlySharedWith.querySelectorAll('[data-sharer-id]')).map(el => el.dataset.sharerId); const availableSharers = sharers.filter(sharer => !alreadySelectedSharerIds.includes(sharer.id) && (sharer.fullname.toLowerCase().includes(lowerSearchTerm) || sharer.email.toLowerCase().includes(lowerSearchTerm))).slice(0, 10); if (availableSharers.length === 0 && searchTerm) { const li = document.createElement('div'); li.classList.add('p-2', 'text-sm', 'text-muted-text'); li.textContent = 'No sharers found.'; modalFileShareSearchResults.appendChild(li); } else { availableSharers.forEach(sharer => { const li = document.createElement('div'); li.classList.add('p-2', 'text-sm', 'cursor-pointer', 'hover:bg-indigo-50'); li.textContent = `${sharer.fullname} (${sharer.email})`; if (sharer.stopAllShares) { li.textContent += ' (Shares Stopped)'; li.classList.add('text-red-500', 'italic'); li.title = "User has sharing stopped."; } li.onclick = () => { if (sharer.stopAllShares && !confirm("User has 'Stop All Shares'. Add anyway?")) return; addFileSharerToSelectedUI(sharer); modalFileShareSearch.value = ''; renderFileShareSearchResults(''); modalFileShareSearch.focus(); }; modalFileShareSearchResults.appendChild(li); }); } }
function addFileSharerToSelectedUI(sharer, expiryDate = null) { modalFileNoSharersMsg.style.display = 'none'; const sharedItemDiv = document.createElement('div'); sharedItemDiv.dataset.sharerId = sharer.id; sharedItemDiv.classList.add('flex', 'items-center', 'justify-between', 'p-2', 'bg-gray-50', 'rounded-md', 'border'); const nameSpan = document.createElement('span'); nameSpan.classList.add('text-sm', 'font-medium', 'text-gray-700', 'truncate', 'max-w-[50%]'); nameSpan.textContent = `${sharer.fullname} (${sharer.email})`; if (sharer.stopAllShares) { nameSpan.textContent += ' (Stopped)'; nameSpan.classList.add('text-red-400'); } sharedItemDiv.appendChild(nameSpan); const controlsDiv = document.createElement('div'); controlsDiv.classList.add('flex', 'items-center', 'space-x-2'); const expiryInput = document.createElement('input'); expiryInput.type = 'datetime-local'; expiryInput.classList.add('p-1.5', 'border', 'border-gray-300', 'rounded-md', 'text-xs', 'focus:ring-brand-blue', 'focus:border-brand-blue'); expiryInput.style.width = '190px'; if (expiryDate) { try { expiryInput.value = new Date(expiryDate).toISOString().slice(0, 16); } catch (e) { console.error("Err formatting date:", expiryDate, e); } } controlsDiv.appendChild(expiryInput); const removeButton = document.createElement('button'); removeButton.textContent = 'Remove'; removeButton.type = 'button'; removeButton.classList.add('text-brand-red', 'hover:text-red-700', 'text-xs', 'p-1', 'font-medium'); removeButton.onclick = () => { sharedItemDiv.remove(); if (modalFileCurrentlySharedWith.children.length === 0 || (modalFileCurrentlySharedWith.children.length === 1 && modalFileCurrentlySharedWith.firstChild.id === 'modalFileNoSharersMsg')) modalFileNoSharersMsg.style.display = 'block'; renderFileShareSearchResults(modalFileShareSearch.value); }; controlsDiv.appendChild(removeButton); sharedItemDiv.appendChild(controlsDiv); modalFileCurrentlySharedWith.appendChild(sharedItemDiv); }
function saveFileShares(fileId) { const newSharesForFile = []; const selectedSharerElements = modalFileCurrentlySharedWith.querySelectorAll('[data-sharer-id]'); const now = new Date().toISOString(); selectedSharerElements.forEach(el => { const sharerId = el.dataset.sharerId; const expiryInput = el.querySelector('input[type="datetime-local"]'); const expiryDate = expiryInput.value ? new Date(expiryInput.value).toISOString() : null; newSharesForFile.push({ id: 'share_' + generateUUID(), itemId: fileId, itemType: 'file', sharerId: sharerId, expiryDate: expiryDate, sharedDate: now }); }); shares = shares.filter(s => !(s.itemId === fileId && s.itemType === 'file')); shares.push(...newSharesForFile); currentFileModalShares = []; }
function renderFolderShareControls(folderId) { currentFolderModalShares = shares.filter(s => s.itemId === folderId && s.itemType === 'folder'); modalFolderCurrentlySharedWith.innerHTML = ''; modalFolderShareSearch.value = ''; renderFolderShareSearchResults(''); modalFolderNoSharersMsg.style.display = currentFolderModalShares.length === 0 ? 'block' : 'none'; currentFolderModalShares.forEach(share => { const sharer = sharers.find(s => s.id === share.sharerId); if (sharer) { addFolderSharerToSelectedUI(sharer, share.expiryDate); } }); }
function renderFolderShareSearchResults(searchTerm) { modalFolderShareSearchResults.innerHTML = ''; const lowerSearchTerm = searchTerm.toLowerCase(); const alreadySelectedSharerIds = Array.from(modalFolderCurrentlySharedWith.querySelectorAll('[data-sharer-id]')).map(el => el.dataset.sharerId); const availableSharers = sharers.filter(sharer => !alreadySelectedSharerIds.includes(sharer.id) && (sharer.fullname.toLowerCase().includes(lowerSearchTerm) || sharer.email.toLowerCase().includes(lowerSearchTerm))).slice(0, 10); if (availableSharers.length === 0 && searchTerm) { const li = document.createElement('div'); li.classList.add('p-2', 'text-sm', 'text-muted-text'); li.textContent = 'No sharers found.'; modalFolderShareSearchResults.appendChild(li); } else { availableSharers.forEach(sharer => { const li = document.createElement('div'); li.classList.add('p-2', 'text-sm', 'cursor-pointer', 'hover:bg-indigo-50'); li.textContent = `${sharer.fullname} (${sharer.email})`; if (sharer.stopAllShares) { li.textContent += ' (Shares Stopped)'; li.classList.add('text-red-500', 'italic'); li.title = "This user currently has all sharing stopped."; } li.onclick = () => { if (sharer.stopAllShares && !confirm("This user has 'Stop All Shares' enabled. Are you sure you want to add them? Their access might be restricted.")) { return; } addFolderSharerToSelectedUI(sharer); modalFolderShareSearch.value = ''; renderFolderShareSearchResults(''); modalFolderShareSearch.focus(); }; modalFolderShareSearchResults.appendChild(li); }); } }
function addFolderSharerToSelectedUI(sharer, expiryDate = null) { modalFolderNoSharersMsg.style.display = 'none'; const sharedItemDiv = document.createElement('div'); sharedItemDiv.dataset.sharerId = sharer.id; sharedItemDiv.classList.add('flex', 'items-center', 'justify-between', 'p-2', 'bg-gray-50', 'rounded-md', 'border'); const nameSpan = document.createElement('span'); nameSpan.classList.add('text-sm', 'font-medium', 'text-gray-700', 'truncate', 'max-w-[50%]'); nameSpan.textContent = `${sharer.fullname} (${sharer.email})`; if (sharer.stopAllShares) { nameSpan.textContent += ' (Stopped)'; nameSpan.classList.add('text-red-400'); } sharedItemDiv.appendChild(nameSpan); const controlsDiv = document.createElement('div'); controlsDiv.classList.add('flex', 'items-center', 'space-x-2'); const expiryInput = document.createElement('input'); expiryInput.type = 'datetime-local'; expiryInput.classList.add('p-1.5', 'border', 'border-gray-300', 'rounded-md', 'text-xs', 'focus:ring-brand-blue', 'focus:border-brand-blue'); expiryInput.style.width = '190px'; if (expiryDate) { try { expiryInput.value = new Date(expiryDate).toISOString().slice(0, 16); } catch (e) { console.error("Error formatting expiry date for folder share:", expiryDate, e); } } controlsDiv.appendChild(expiryInput); const removeButton = document.createElement('button'); removeButton.textContent = 'Remove'; removeButton.type = 'button'; removeButton.classList.add('text-brand-red', 'hover:text-red-700', 'text-xs', 'p-1', 'font-medium'); removeButton.onclick = () => { sharedItemDiv.remove(); if (modalFolderCurrentlySharedWith.children.length === 0 || (modalFolderCurrentlySharedWith.children.length === 1 && modalFolderCurrentlySharedWith.firstChild.id === 'modalFolderNoSharersMsg')) { modalFolderNoSharersMsg.style.display = 'block'; } renderFolderShareSearchResults(modalFolderShareSearch.value); }; controlsDiv.appendChild(removeButton); sharedItemDiv.appendChild(controlsDiv); modalFolderCurrentlySharedWith.appendChild(sharedItemDiv); }
function saveFolderShares(folderId) { const newSharesForFolder = []; const selectedSharerElements = modalFolderCurrentlySharedWith.querySelectorAll('[data-sharer-id]'); const now = new Date().toISOString(); selectedSharerElements.forEach(el => { const sharerId = el.dataset.sharerId; const expiryInput = el.querySelector('input[type="datetime-local"]'); const expiryDate = expiryInput.value ? new Date(expiryInput.value).toISOString() : null; newSharesForFolder.push({ id: 'share_' + generateUUID(), itemId: folderId, itemType: 'folder', sharerId: sharerId, expiryDate: expiryDate, sharedDate: now }); }); shares = shares.filter(s => !(s.itemId === folderId && s.itemType === 'folder')); shares.push(...newSharesForFolder); currentFolderModalShares = []; }

function openViewFileModal(fileIdToView) {
    const file = files.find(f => f.id === fileIdToView && !f.isDeleted);
    if (!file) { alert("File not found or cannot be viewed."); return; }
    if (!file.driveFileId) { alert("This file is not synced to Google Drive and cannot be previewed."); return; }
    viewFileModalTitle.textContent = `Viewing: ${file.currentName}`;
    const iframeSrc = `https://drive.google.com/file/d/${file.driveFileId}/preview`;
    viewFileIframe.src = iframeSrc;
    viewFileModal.classList.remove('hidden');
}
function closeViewFileModal() { viewFileModal.classList.add('hidden'); viewFileIframe.src = 'about:blank'; viewFileModalTitle.textContent = 'Viewing File...'; }


// --- MODIFIED Deep Link Processing ---
async function checkAndProcessDeepLink() {
    console.log("Checking for deep links... Current auth state: GDrive Ready=", isDriveReadyForOps, "MS Authenticated=", isMicrosoftAuthenticated);

    const deepLinkedFileId = sessionStorage.getItem('urlQueryFileId');
    const deepLinkedFolderId = sessionStorage.getItem('urlQueryFolderId');

    const localDataExists = localStorage.getItem(LOCAL_STORAGE_INDEX_KEY) !== null;
    const dataIsSufficientlyLoaded = files.length > 0 || folders.length > 0 || isDriveReadyForOps || localDataExists;

    if (!dataIsSufficientlyLoaded && (deepLinkedFileId || deepLinkedFolderId)) {
        console.log("Deep link: Data (files/folders) not sufficiently loaded yet. Deferring processing of deep link items.");
        deepLinkProcessed = false;
        return;
    }

    if (deepLinkedFileId) {
        console.log(`Attempting to process deep link for file ID: ${deepLinkedFileId}`);
        const fileToOpen = files.find(f => f.id === deepLinkedFileId && !f.isDeleted);

        if (fileToOpen) {
            if (fileToOpen.encryptedDriveJsonId) { // Encrypted file
                // For encrypted files, both GDrive (for JSON) and MS (for fragment & password) must be ready
                if (isDriveReadyForOps && isMicrosoftAuthenticated) {
                    console.log("Deep link: Auth complete for encrypted file. Attempting to open encrypted preview for file:", fileToOpen.currentName);
                    await fetchAndCacheEncryptedJsonForPreview(fileToOpen.encryptedDriveJsonId, fileToOpen.id, fileToOpen.currentName);
                    sessionStorage.removeItem('urlQueryFileId');
                } else {
                    console.warn(`Deep link: Encrypted file "${fileToOpen.currentName}" preview requires GDrive & MS auth. Auth not yet fully complete. (GDrive: ${isDriveReadyForOps}, MS: ${isMicrosoftAuthenticated}). Will retry when auth state changes.`);
                }
            } else { // Non-encrypted file
                console.log("Deep link: Non-encrypted file ID found, opening metadata modal for:", fileToOpen.currentName);
                openMetadataModal(deepLinkedFileId);
                sessionStorage.removeItem('urlQueryFileId');
            }
        } else {
            if (dataIsSufficientlyLoaded) {
                console.warn("Deep link: File ID from URL not found or deleted:", deepLinkedFileId);
                alert(`The linked file (ID: ${deepLinkedFileId}) was not found or has been deleted.`);
                sessionStorage.removeItem('urlQueryFileId');
            } else {
                console.log("Deep link (file): Data not loaded enough to confirm 'not found' for file ID:", deepLinkedFileId);
            }
        }
    }

    if (deepLinkedFolderId) {
        console.log(`Attempting to process deep link for folder ID: ${deepLinkedFolderId}`);
        const folderToOpen = folders.find(f => f.id === deepLinkedFolderId && !f.isDeleted);
        if (folderToOpen) {
            setActiveFolder(deepLinkedFolderId);
            if (activeFolderId === deepLinkedFolderId) {
                const pathChanged = ensurePathExpanded(deepLinkedFolderId);
                if (pathChanged) _commitLocalStateChanges();
                renderAll();
            }
            sessionStorage.removeItem('urlQueryFolderId');
        } else {
            if (dataIsSufficientlyLoaded) {
                console.warn("Deep link: Folder ID from URL not found or deleted:", deepLinkedFolderId);
                alert(`The linked folder (ID: ${deepLinkedFolderId}) was not found or has been deleted.`);
                sessionStorage.removeItem('urlQueryFolderId');
            } else {
                console.log("Deep link (folder): Data not loaded enough to confirm 'not found' for folder ID:", deepLinkedFolderId);
            }
        }
    }

    if (!sessionStorage.getItem('urlQueryFileId') && !sessionStorage.getItem('urlQueryFolderId')) {
        deepLinkProcessed = true;
        console.log("Deep link: All session items processed/cleared. Marking deepLinkProcessed = true.");
    } else {
        deepLinkProcessed = false;
        console.log("Deep link: Items still in session storage. Marking deepLinkProcessed = false for future re-checks.");
    }
    console.log("Deep link check complete. Current deepLinkProcessed flag:", deepLinkProcessed, "Session FileID:", sessionStorage.getItem('urlQueryFileId'), "Session FolderID:", sessionStorage.getItem('urlQueryFolderId'));
}


// --- Initialization Page Load ---
// MODIFIED: Now initializes IndexedDB on page load.
window.onload = async () => {
    // ADDED: Initialize the IndexedDB helper.
    await dbHelper.init().catch(err => {
        console.error("Failed to initialize IndexedDB. Caching will not work.", err);
        alert("CRITICAL: Could not initialize local database for caching. Encrypted file previews will not work. Please check browser settings (e.g., allow data storage).");
    });

    try {
        const urlParams = new URLSearchParams(window.location.search);
        const fileIdFromUrl = urlParams.get('fileid');
        const folderIdFromUrl = urlParams.get('folderid');

        if (fileIdFromUrl) sessionStorage.setItem('urlQueryFileId', fileIdFromUrl);
        if (folderIdFromUrl) sessionStorage.setItem('urlQueryFolderId', folderIdFromUrl);
        if (fileIdFromUrl || folderIdFromUrl) {
            const newUrlParams = new URLSearchParams();
            urlParams.forEach((value, key) => {
                if (key.toLowerCase() !== 'fileid' && key.toLowerCase() !== 'folderid') newUrlParams.append(key, value);
            });
            history.replaceState(null, '', `${window.location.pathname}${newUrlParams.toString() ? '?' + newUrlParams.toString() : ''}`);
        }
    } catch (e) { console.error("Error processing URL parameters:", e); }

    filesDataTableHeaders = document.querySelectorAll('#contentDataTable th.sortable-header');
    sharersDataTableHeaders = document.querySelectorAll('#contentSharers th.sortable-header');

    loadDataFromLocalStorage();
    renderAll();

    if (microsoftSignInBtn && microsoftApiStatusEl) {
        microsoftSignInBtn.addEventListener('click', msLogin);
        try {
            await msalInstance.initialize();
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                msalInstance.setActiveAccount(accounts[0]);
                microsoftAccount = accounts[0];
                isMicrosoftAuthenticated = true;
                microsoftApiStatusEl.textContent = `✅ Auto signed in to Microsoft as ${microsoftAccount.username}.`;
                microsoftSignInBtn.textContent = `✅ MS: ${microsoftAccount.name || microsoftAccount.username.split('@')[0]}`;
                microsoftSignInBtn.disabled = true;
                const token = await getMicrosoftAccessToken();
                if (token) {
                    await ensureOneDriveFolderExists(token, ONEDRIVE_FRAGMENT_FOLDER_NAME);
                    if (oneDriveFragmentFolderId) microsoftApiStatusEl.textContent += ` Using OneDrive folder: "${ONEDRIVE_FRAGMENT_FOLDER_NAME}".`;
                    else microsoftApiStatusEl.textContent += ` Error ensuring OneDrive folder. Check console.`;
                }
            } else {
                microsoftApiStatusEl.textContent = "Please sign in to Microsoft for fragment & password storage.";
            }
        } catch (msalInitError) {
            console.error("MSAL Initialization Error:", msalInitError);
            microsoftApiStatusEl.textContent = "❌ MSAL Init Failed. Check console.";
        } finally {
            await checkAndProcessDeepLink();
        }
    } else {
        console.error("Microsoft Sign In Button or Status Element not found.");
        await checkAndProcessDeepLink();
    }

    if (googleSignInBtn && googleApiStatusEl) {
        try {
            const savedToken = localStorage.getItem('googleAccessToken');
            const savedTimestamp = parseInt(localStorage.getItem('googleAccessTokenTimestamp'), 10);
            const tokenAge = Date.now() - savedTimestamp;
            const TOKEN_VALID_MS = 60 * 60 * 1000;

            if (savedToken && tokenAge < TOKEN_VALID_MS) {
                googleAccessToken = savedToken;
                isDriveAuthenticated = true;
                console.log("✅ Restored Google token from localStorage. Initializing GAPI client...");
                gapi.load('client', initializeGapiClient);
            } else {
                localStorage.removeItem('googleAccessToken');
                localStorage.removeItem('googleAccessTokenTimestamp');
                if (googleApiStatusEl) googleApiStatusEl.textContent = "Please sign in to sync with Google Drive. Displaying local data if available.";
                await checkAndProcessDeepLink();
            }
        } catch (e) {
            console.error("Failed to initialize Google Token Client or process token", e);
            if (googleApiStatusEl) googleApiStatusEl.textContent = "Could not initialize Google Sign-In. Check console.";
            await checkAndProcessDeepLink();
        }
        googleSignInBtn.addEventListener('click', () => { authenticateGoogle() }); // Ensure it's called with default false
    } else {
        console.error("Google Sign In Button or Status Element not found.");
        await checkAndProcessDeepLink();
    }

    // Event Listeners (condensed for brevity, ensure they are complete)
    if (dropZone) { dropZone.addEventListener('click', () => fileInput.click()); dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('border-brand-blue', 'bg-indigo-50'); }); dropZone.addEventListener('dragleave', (e) => { e.preventDefault(); dropZone.classList.remove('border-brand-blue', 'bg-indigo-50'); }); dropZone.addEventListener('drop', (e) => { e.preventDefault(); dropZone.classList.remove('border-brand-blue', 'bg-indigo-50'); handleFileProcessingAndDriveUpload(e.dataTransfer.files); }); }
    if (fileInput) fileInput.addEventListener('change', (e) => { handleFileProcessingAndDriveUpload(e.target.files); fileInput.value = ''; });
    if (newFolderBtn) newFolderBtn.addEventListener('click', openNewFolderModal);
    if (modalNewFolderCancelBtn) modalNewFolderCancelBtn.addEventListener('click', () => newFolderModal.classList.add('hidden'));
    if (modalRenameFolderCancelBtn) modalRenameFolderCancelBtn.addEventListener('click', () => renameFolderModal.classList.add('hidden'));
    if (modalCancelBtn) modalCancelBtn.addEventListener('click', () => metadataModal.classList.add('hidden'));
    if (tabFileTreeBtn) tabFileTreeBtn.addEventListener('click', () => switchTab('fileTree'));
    if (tabDataTableBtn) tabDataTableBtn.addEventListener('click', () => switchTab('dataTable'));
    if (tabSharersBtn) tabSharersBtn.addEventListener('click', () => switchTab('sharers'));
    if (searchInput) searchInput.addEventListener('input', () => { filesDataTablePage = 1; renderDataTable(searchInput.value); });
    if (itemsPerPageFilesSelect) itemsPerPageFilesSelect.addEventListener('change', (e) => { itemsPerPageFiles = parseInt(e.target.value); localStorage.setItem('itemsPerPageFiles', itemsPerPageFiles); filesDataTablePage = 1; renderDataTable(searchInput.value); });
    if (prevPageBtnFiles) prevPageBtnFiles.addEventListener('click', () => { if (filesDataTablePage > 1) { filesDataTablePage--; renderDataTable(searchInput.value); } });
    if (nextPageBtnFiles) nextPageBtnFiles.addEventListener('click', () => { /* ... (logic from original) ... */ const lowerSearchTerm = searchInput ? searchInput.value.toLowerCase() : ''; let displayFiles = files.filter(file => !file.isDeleted); if (lowerSearchTerm) { /* filter logic */ displayFiles = displayFiles.filter(file => { const folder = folders.find(f => f.id === file.folderId && !f.isDeleted); const folderName = folder ? folder.name : (file.folderId === 'root' ? 'Root' : 'N/A'); return file.currentName.toLowerCase().includes(lowerSearchTerm) || file.originalName.toLowerCase().includes(lowerSearchTerm) || folderName.toLowerCase().includes(lowerSearchTerm) || (file.tags && file.tags.join(',').toLowerCase().includes(lowerSearchTerm)) || (file.comments && file.comments.toLowerCase().includes(lowerSearchTerm)) || file.mimeType.toLowerCase().includes(lowerSearchTerm); }); } const totalItems = displayFiles.length; const totalPages = Math.max(1, Math.ceil(totalItems / itemsPerPageFiles)); if (filesDataTablePage < totalPages) { filesDataTablePage++; renderDataTable(searchInput.value); } });
    if (filesDataTableHeaders) filesDataTableHeaders.forEach(header => header.addEventListener('click', () => { const sortKey = header.dataset.sort; if (filesSortColumn === sortKey) filesSortDirection = filesSortDirection === 'asc' ? 'desc' : 'asc'; else { filesSortColumn = sortKey; filesSortDirection = 'asc'; } renderDataTable(searchInput.value); }));
    if (sharersSearchInput) sharersSearchInput.addEventListener('input', () => { sharersTablePage = 1; renderSharersTab(sharersSearchInput.value); });
    if (itemsPerPageSharersSelect) itemsPerPageSharersSelect.addEventListener('change', (e) => { itemsPerPageSharers = parseInt(e.target.value); localStorage.setItem('itemsPerPageSharers', itemsPerPageSharers); sharersTablePage = 1; renderSharersTab(sharersSearchInput.value); });
    if (prevPageBtnSharers) prevPageBtnSharers.addEventListener('click', () => { if (sharersTablePage > 1) { sharersTablePage--; renderSharersTab(sharersSearchInput.value); } });
    if (nextPageBtnSharers) nextPageBtnSharers.addEventListener('click', () => { /* ... (logic from original) ... */ const lowerSearchTerm = sharersSearchInput ? sharersSearchInput.value.toLowerCase() : ''; let displaySharers = [...sharers]; if (lowerSearchTerm) { /* filter logic */ displaySharers = displaySharers.filter(sharer => sharer.shortname.toLowerCase().includes(lowerSearchTerm) || sharer.fullname.toLowerCase().includes(lowerSearchTerm) || sharer.email.toLowerCase().includes(lowerSearchTerm) || (sharer.dept && sharer.dept.toLowerCase().includes(lowerSearchTerm))); } const totalItems = displaySharers.length; const totalPages = Math.max(1, Math.ceil(totalItems / itemsPerPageSharers)); if (sharersTablePage < totalPages) { sharersTablePage++; renderSharersTab(sharersSearchInput.value); } });
    if (sharersDataTableHeaders) sharersDataTableHeaders.forEach(header => header.addEventListener('click', () => { const sortKey = header.dataset.sortSharers; if (sharersSortColumn === sortKey) sharersSortDirection = sharersSortDirection === 'asc' ? 'desc' : 'asc'; else { sharersSortColumn = sortKey; filesSortDirection = 'asc'; } renderSharersTab(sharersSearchInput.value); })); // Corrected: filesSortDirection to sharersSortDirection
    if (newSharerBtn) newSharerBtn.addEventListener('click', () => openSharerModal());
    if (modalSharerCancelBtn) modalSharerCancelBtn.addEventListener('click', () => sharerModal.classList.add('hidden'));
    if (modalFileShareSearch) modalFileShareSearch.addEventListener('input', (e) => renderFileShareSearchResults(e.target.value));
    if (modalFileShareSearch) modalFileShareSearch.addEventListener('focus', () => { if (!modalFileShareSearch.value) renderFileShareSearchResults(''); });
    if (modalFolderShareSearch) modalFolderShareSearch.addEventListener('input', (e) => renderFolderShareSearchResults(e.target.value));
    if (modalFolderShareSearch) modalFolderShareSearch.addEventListener('focus', () => { if (!modalFolderShareSearch.value) renderFolderShareSearchResults(''); });
    document.addEventListener('click', function (event) { if (modalFileShareSearch && modalFileShareSearchResults && !modalFileShareSearch.contains(event.target) && !modalFileShareSearchResults.contains(event.target) && !event.target.closest('#modalFileShareSearchResults')) { modalFileShareSearchResults.innerHTML = ''; } if (modalFolderShareSearch && modalFolderShareSearchResults && !modalFolderShareSearch.contains(event.target) && !modalFolderShareSearchResults.contains(event.target) && !event.target.closest('#modalFolderShareSearchResults')) { modalFolderShareSearchResults.innerHTML = ''; } });

    // Ensure UPLOAD_BUTTON_SVG_ICON and UPLOAD_BUTTON_TEXT_SPAN_ID are defined if used, or handle missing elements gracefully
    const UPLOAD_BUTTON_SVG_ICON = `<svg xmlns="http://www.w3.org/2000/svg" class="h-5 w-5 mr-2 pointer-events-none" viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zM6.293 6.707a1 1 0 010-1.414l3-3a1 1 0 011.414 0l3 3a1 1 0 01-1.414 1.414L11 5.414V13a1 1 0 11-2 0V5.414L7.707 6.707a1 1 0 01-1.414 0z" clip-rule="evenodd" /></svg>`;
    const UPLOAD_BUTTON_TEXT_SPAN_ID = 'uploadButtonTextSpan'; // Example ID

    if (toggleUploadAreaBtn && uploadAreaWrapper) {
        // Initialize button text based on initial state
        const isUploadAreaHidden = uploadAreaWrapper.classList.contains('hidden');
        let initialUploadButtonText = isUploadAreaHidden ? '+ Upload' : '- Upload';
        let textSpanToUpdate = toggleUploadAreaBtn.querySelector('.sidebar-text'); // General selector
        if (!textSpanToUpdate && document.getElementById(UPLOAD_BUTTON_TEXT_SPAN_ID)) { // Fallback to ID if general selector fails
            textSpanToUpdate = document.getElementById(UPLOAD_BUTTON_TEXT_SPAN_ID);
        }
        if (textSpanToUpdate) { // If span exists, update its text
            textSpanToUpdate.textContent = initialUploadButtonText;
        } else { // If span doesn't exist, reconstruct carefully
            toggleUploadAreaBtn.innerHTML = UPLOAD_BUTTON_SVG_ICON + `<span class="sidebar-text pointer-events-none" ${document.getElementById(UPLOAD_BUTTON_TEXT_SPAN_ID) ? 'id="' + UPLOAD_BUTTON_TEXT_SPAN_ID + '"' : ''}>${initialUploadButtonText}</span>`;
        }


        toggleUploadAreaBtn.addEventListener('click', () => {
            uploadAreaWrapper.classList.toggle('hidden');
            const isHidden = uploadAreaWrapper.classList.contains('hidden');
            const newText = isHidden ? '+ Upload' : '- Upload';

            let textSpan = toggleUploadAreaBtn.querySelector('.sidebar-text');
            if (!textSpan && document.getElementById(UPLOAD_BUTTON_TEXT_SPAN_ID)) {
                textSpan = document.getElementById(UPLOAD_BUTTON_TEXT_SPAN_ID);
            }

            if (textSpan) {
                textSpan.textContent = newText;
            } else { // Fallback if span somehow got removed - reconstruct
                toggleUploadAreaBtn.innerHTML = UPLOAD_BUTTON_SVG_ICON + `<span class="sidebar-text pointer-events-none" ${document.getElementById(UPLOAD_BUTTON_TEXT_SPAN_ID) ? 'id="' + UPLOAD_BUTTON_TEXT_SPAN_ID + '"' : ''}>${newText}</span>`;
            }
        });
    }
    if (viewFileModalCloseBtn) viewFileModalCloseBtn.addEventListener('click', closeViewFileModal);
    if (viewFileModal) viewFileModal.addEventListener('click', (event) => { if (event.target === viewFileModal) closeViewFileModal(); });
};

// --- MODIFIED: Logout routine with robust IndexedDB clearing ---
document.addEventListener('DOMContentLoaded', () => {
    const logoutButton = document.getElementById('logout');
    if (logoutButton) {
        logoutButton.addEventListener('click', async () => {
            try {
                // Step 1: Close our current DB connection to allow deletion.
                if (dbHelper.db) {
                    dbHelper.db.close();
                    dbHelper.db = null;
                    console.log("IndexedDB connection closed for logout.");
                }

                // Step 2: Try to delete the entire database.
                console.log(`Attempting to delete IndexedDB: ${dbHelper.dbName}`);
                await new Promise((resolve, reject) => {
                    const deleteRequest = indexedDB.deleteDatabase(dbHelper.dbName);
                    deleteRequest.onsuccess = () => {
                        console.log("IndexedDB database deleted successfully.");
                        resolve();
                    };
                    deleteRequest.onerror = (e) => {
                        console.error("Error deleting IndexedDB.", e);
                        reject(new Error("DB deletion failed."));
                    };
                    deleteRequest.onblocked = (e) => {
                        // This is a critical event. It means another tab has the DB open.
                        console.warn("IndexedDB deletion blocked. Another tab may be open.", e);
                        reject(new Error("DB deletion blocked by another tab."));
                    };
                });
            } catch (error) {
                // Step 3 (Fallback): If deletion fails (e.g., it's blocked), open it and clear the store.
                console.warn(`Could not delete IndexedDB (${error.message}). Attempting to clear the store as a fallback.`);
                try {
                    await dbHelper.init(); // Re-open a connection
                    const transaction = dbHelper.db.transaction(dbHelper.storeName, 'readwrite');
                    const store = transaction.objectStore(dbHelper.storeName);
                    await new Promise((resolve, reject) => {
                        const clearRequest = store.clear();
                        clearRequest.onsuccess = () => {
                            console.log("IndexedDB object store cleared successfully.");
                            resolve();
                        };
                        clearRequest.onerror = (e) => {
                            console.error("Error clearing object store.", e);
                            reject(e);
                        };
                    });
                } catch (clearError) {
                    console.error("CRITICAL: Failed to both delete and clear IndexedDB.", clearError);
                    alert("Warning: Could not fully clear all cached data. For maximum security, please consider manually clearing your browser's site data for this page.");
                }
            } finally {
                // Step 4: Always perform the rest of the logout actions.
                console.log("Clearing localStorage and sessionStorage.");
                localStorage.clear();
                sessionStorage.clear();

                if (msalInstance && msalInstance.getActiveAccount()) {
                    msalInstance.logoutPopup({ mainWindowRedirectUri: window.location.origin })
                        .finally(() => { // Use finally to ensure reload happens even if logout fails
                            location.reload();
                        });
                } else {
                    location.reload();
                }
            }
        });
    }
    const menuToggle = document.getElementById('menu-toggle');
    const menu = document.getElementById('menu');
    if (menuToggle && menu) menuToggle.addEventListener('click', () => menu.classList.toggle('hidden'));
});

setInterval(() => {
    if (tokenClient && isDriveAuthenticated) {
        console.log("Attempting silent Google token refresh...");
        tokenClient.requestAccessToken({ prompt: '' });
    }
}, 50 * 60 * 1000);
// --- END OF FULL main3.js ---
// --- END OF FILE main3.js ---
// --- END OF FILE main3.js ---