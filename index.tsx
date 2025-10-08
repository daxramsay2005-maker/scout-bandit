/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */
import * as L from 'leaflet';
import { GoogleGenAI, Type } from '@google/genai';

// --- Type Declarations for Google APIs ---
declare const gapi: any;
declare const google: any;

// --- DOM Elements ---
const businessTypeInput = document.getElementById('business-type-input') as HTMLInputElement;
const cityInput = document.getElementById('city-input') as HTMLInputElement;
const stateInput = document.getElementById('state-input') as HTMLInputElement;
const searchRadiusInput = document.getElementById('search-radius-input') as HTMLInputElement;
const aiSearchButton = document.getElementById('ai-search-button') as HTMLButtonElement;
const exportCsvButton = document.getElementById('export-csv-button') as HTMLButtonElement;
const resultsContainer = document.getElementById('results-container') as HTMLDivElement;
const mapContainer = document.getElementById('map') as HTMLDivElement;
const searchInput = document.getElementById('search-input') as HTMLInputElement;
const instagramHandleInput = document.getElementById('instagram-handle-input') as HTMLInputElement;
const followsFilterCheckbox = document.getElementById('follows-filter-checkbox') as HTMLInputElement;
const authContainer = document.getElementById('auth-container') as HTMLDivElement;
const authErrorContainer = document.getElementById('auth-error-message') as HTMLDivElement;
const googleClientIdInput = document.getElementById('google-client-id-input') as HTMLInputElement;
const saveToSheetsButton = document.getElementById('save-to-sheets-button') as HTMLButtonElement;
const saveToDriveButton = document.getElementById('save-to-drive-button') as HTMLButtonElement;
const statusMessage = document.getElementById('status-message') as HTMLDivElement;

// Modal Elements
const detailsModal = document.getElementById('details-modal') as HTMLDivElement;
const modalBusinessName = document.getElementById('modal-business-name') as HTMLHeadingElement;
const modalBody = document.getElementById('modal-body') as HTMLDivElement;
const modalCloseButton = document.getElementById('modal-close-button') as HTMLButtonElement;


// --- Google API Config ---
const API_KEY = process.env.API_KEY;
const ai = new GoogleGenAI({ apiKey: API_KEY });

const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file';


// --- App State ---
let map: L.Map | null = null;
let markersLayer: L.FeatureGroup | null = null;
let currentBusinesses: any[] = [];
let tokenClient: any = null;
let gapiInited = false;
let gisInited = false; // True once the GIS script has loaded

// --- Utility Functions ---

function escapeHTML(str: any): string {
    if (typeof str !== 'string') return String(str || '');
    return str.replace(/[&<>"']/g, (match) => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;'
    }[match]!));
}

function showStatus(message: string, type: 'success' | 'error', duration: number = 5000) {
    statusMessage.innerHTML = message;
    statusMessage.className = `visible ${type}`;
    if (duration > 0) {
        setTimeout(() => {
            statusMessage.className = '';
        }, duration);
    }
}

function setButtonLoadingState(button: HTMLButtonElement, isLoading: boolean) {
    button.disabled = isLoading;
    if (isLoading) {
        button.classList.add('loading');
    } else {
        button.classList.remove('loading');
    }
}

// --- Data Handling ---
const getBusinessId = (business: any): string => btoa(encodeURIComponent(business.name + business.address));

const convertToCsv = (data: any[]): string => {
    if (data.length === 0) return '';
    const headers = Object.keys(data[0]).filter(key => key !== 'socialMedia');
    const socialHeaders = data.some(d => d.socialMedia) ? Object.keys(data.find(d => d.socialMedia)?.socialMedia || {}) : [];
    const allHeaders = [...headers, ...socialHeaders];

    const headerRow = allHeaders.map(h => `"${h}"`).join(',');
    const rows = data.map(row => {
        const mainValues = headers.map(header => `"${escapeHTML(row[header])}"`);
        const socialValues = socialHeaders.map(header => `"${escapeHTML(row.socialMedia?.[header] || '')}"`);
        return [...mainValues, ...socialValues].join(',');
    });
    return [headerRow, ...rows].join('\n');
}

// --- Google Auth ---

async function initializeGapiClient() {
    try {
        await gapi.client.init({});
        await gapi.client.load('https://sheets.googleapis.com/$discovery/rest?version=v4');
        await gapi.client.load('https://www.googleapis.com/discovery/v1/apis/drive/v3/rest');
        gapiInited = true;
        maybeEnableAuth();
    } catch (error) {
        console.error("Error initializing GAPI client:", error);
        showStatus('Could not connect to Google services.', 'error');
    }
}

function initializeGisClient(clientId: string) {
    try {
        tokenClient = google.accounts.oauth2.initTokenClient({
            client_id: clientId,
            scope: SCOPES,
            callback: (tokenResponse: any) => {
                gapi.client.setToken(tokenResponse);
                updateAuthUI(true);
            },
            error_callback: (error: any) => {
                console.error("Google Sign-In Error:", error);
                const signInButton = document.getElementById('authorize-button') as HTMLButtonElement;
                
                // Reset button and token client
                if (signInButton) setButtonLoadingState(signInButton, false);
                tokenClient = null;

                let errorMessage = '';
                let transientMessage = '';

                if (error.type === 'popup_closed_by_user') {
                    transientMessage = 'Sign-in cancelled.';
                } else if (error.type === 'popup_failed_to_open') {
                    errorMessage = 'Sign-in failed. Please disable your popup blocker and try again.';
                } else if (error.type === 'token_request_failed' || error.type === 'idpiframe_initialization_failed' || error.type === 'unknown') {
                    errorMessage = `
                        <strong>Authentication Configuration Error</strong><br>
                        This is likely due to a misconfiguration in your Google Cloud project. Please ensure this app's URL is added to your list of <b>Authorized JavaScript origins</b>.
                        <br><br>
                        <a href="https://console.cloud.google.com/apis/credentials" target="_blank" rel="noopener noreferrer">Go to Google Cloud Credentials to fix this.</a>
                    `;
                } else {
                    errorMessage = `An unexpected sign-in error occurred: ${error.type || 'Unknown'}.`;
                }
                
                if (transientMessage) {
                     showStatus(transientMessage, 'success');
                }

                if (errorMessage) {
                    authErrorContainer.innerHTML = errorMessage;
                    authErrorContainer.classList.add('visible');
                }
            }
        });
    } catch (error) {
        console.error("Error initializing GIS client:", error);
        showStatus('Could not set up Google Sign-In. Check your Client ID.', 'error');
        tokenClient = null; // Reset on failure
    }
}

function maybeEnableAuth() {
    if (gapiInited && gisInited) {
        updateAuthUI(false); // Assume signed out initially
    }
}

async function updateAuthUI(isAuthed: boolean) {
    authContainer.innerHTML = '';
    authErrorContainer.classList.remove('visible');
    authErrorContainer.innerHTML = '';
    saveToSheetsButton.disabled = !isAuthed;
    saveToDriveButton.disabled = !isAuthed;

    if (isAuthed) {
        try {
            const userInfoRes = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
                headers: { 'Authorization': `Bearer ${gapi.client.getToken().access_token}` }
            });
            if (!userInfoRes.ok) throw new Error('Failed to fetch user info');
            const userInfo = await userInfoRes.json();

            const profileDiv = document.createElement('div');
            profileDiv.className = 'user-profile';
            profileDiv.innerHTML = `<img src="${escapeHTML(userInfo.picture)}" alt="User profile picture" /><span>${escapeHTML(userInfo.given_name)}</span>`;
            authContainer.appendChild(profileDiv);

            const signOutButton = document.createElement('button');
            signOutButton.id = 'signout-button';
            signOutButton.textContent = 'Sign Out';
            signOutButton.addEventListener('click', handleSignoutClick);
            authContainer.appendChild(signOutButton);
        } catch (error) {
            console.error("Error fetching user info:", error);
            const signOutButton = document.createElement('button'); // Fallback
            signOutButton.id = 'signout-button';
            signOutButton.textContent = 'Sign Out';
            signOutButton.addEventListener('click', handleSignoutClick);
            authContainer.appendChild(signOutButton);
        }
    } else {
        const signInButton = document.createElement('button');
        signInButton.id = 'authorize-button';
        signInButton.innerHTML = `<span class="button-text">Sign in with Google</span>`;
        signInButton.addEventListener('click', handleAuthClick);
        signInButton.disabled = !googleClientIdInput.value.trim(); // Disable if no client ID
        authContainer.appendChild(signInButton);
    }
}

function handleAuthClick() {
    // Hide previous errors on a new attempt
    authErrorContainer.classList.remove('visible');
    authErrorContainer.innerHTML = '';

    const clientId = googleClientIdInput.value.trim();
    if (!clientId) {
        showStatus('Please enter your Google Client ID first.', 'error');
        return;
    }

    const signInButton = document.getElementById('authorize-button') as HTMLButtonElement;
    if (signInButton) setButtonLoadingState(signInButton, true);
    
    // Initialize the token client if it hasn't been already
    if (!tokenClient) {
        initializeGisClient(clientId);
    }

    // If initialization failed, tokenClient will be null
    if (!tokenClient) {
        if (signInButton) setButtonLoadingState(signInButton, false);
        return;
    }

    if (gapi.client.getToken() === null) {
        tokenClient.requestAccessToken({ prompt: 'consent' });
    } else {
        tokenClient.requestAccessToken({ prompt: '' });
    }
}

function handleSignoutClick() {
    const token = gapi.client.getToken();
    if (token !== null) {
        google.accounts.oauth2.revoke(token.access_token, () => {
            gapi.client.setToken(null);
            tokenClient = null; // Reset the token client
            updateAuthUI(false);
            showStatus('You have been signed out.', 'success');
        });
    }
}

// --- AI Search ---

async function handleAiSearch() {
    const businessType = businessTypeInput.value.trim();
    const city = cityInput.value.trim();
    const state = stateInput.value.trim();
    const radius = searchRadiusInput.value.trim();

    if (!businessType || !city || !state || !radius) {
        showStatus('Please fill in all search fields.', 'error');
        return;
    }

    setButtonLoadingState(aiSearchButton, true);
    resultsContainer.innerHTML = '<div class="loader"></div>';
    mapContainer.classList.add('hidden');
    exportCsvButton.classList.add('hidden');
    saveToSheetsButton.classList.add('hidden');
    saveToDriveButton.classList.add('hidden');


    const prompt = `Find ${businessType} in or near ${city}, ${state} within a ${radius} mile radius. For each business, provide the name, full address, phone number, website URL, a 1-2 sentence descriptive summary, an estimated customer rating out of 5 (as a number), and a list of social media URLs (like Instagram, Twitter, Facebook).`;
    const responseSchema = {
        type: Type.ARRAY,
        items: {
            type: Type.OBJECT,
            properties: {
                name: { type: Type.STRING },
                address: { type: Type.STRING },
                phone: { type: Type.STRING },
                website: { type: Type.STRING },
                summary: { type: Type.STRING },
                rating: { type: Type.NUMBER },
                lat: { type: Type.NUMBER, description: 'Latitude for mapping.' },
                lng: { type: Type.NUMBER, description: 'Longitude for mapping.' },
                socialMedia: {
                    type: Type.OBJECT,
                    properties: {
                        instagram: { type: Type.STRING },
                        twitter: { type: Type.STRING },
                        facebook: { type: Type.STRING },
                    },
                },
            },
            required: ['name', 'address', 'lat', 'lng']
        }
    };

    try {
        const response = await ai.models.generateContent({
            model: 'gemini-2.5-flash',
            contents: prompt,
            config: {
                responseMimeType: 'application/json',
                responseSchema,
            },
        });
        const businesses = JSON.parse(response.text.trim());
        currentBusinesses = businesses;
        if (businesses.length > 0) {
            initMap();
        }
        filterAndRenderResults();
    } catch (error) {
        console.error('AI Search Error:', error);
        showStatus('Failed to get results from AI. Please try again.', 'error');
        resultsContainer.innerHTML = '<p class="error">An error occurred. Please refine your search and try again.</p>';
    } finally {
        setButtonLoadingState(aiSearchButton, false);
    }
}

// --- Rendering ---
function renderResults(businesses: any[]) {
    resultsContainer.innerHTML = '';
    if (businesses.length === 0) {
        resultsContainer.innerHTML = '<p>No businesses found matching your criteria.</p>';
        mapContainer.classList.add('hidden');
        exportCsvButton.classList.add('hidden');
        saveToSheetsButton.classList.add('hidden');
        saveToDriveButton.classList.add('hidden');
        return;
    }

    businesses.forEach(business => {
        resultsContainer.appendChild(createBusinessCard(business));
    });

    updateMapMarkers(businesses);
    mapContainer.classList.remove('hidden');
    exportCsvButton.classList.remove('hidden');
    saveToSheetsButton.classList.remove('hidden');
    saveToDriveButton.classList.remove('hidden');
}

function createBusinessCard(business: any): HTMLElement {
    const card = document.createElement('div');
    card.className = 'business-card';
    card.dataset.id = getBusinessId(business);

    const socialLinks = business.socialMedia ? Object.entries(business.socialMedia).map(([platform, url]) => {
        if (!url) return '';
        return `<a href="${escapeHTML(url)}" target="_blank" rel="noopener noreferrer" class="social-link" aria-label="${platform}">${platform.charAt(0).toUpperCase() + platform.slice(1)}</a>`;
    }).join('') : 'N/A';

    card.innerHTML = `
        <div class="card-header">
            <h3>${escapeHTML(business.name)}</h3>
        </div>
        ${business.rating ? `<div class="rating">Rating: ${escapeHTML(business.rating)} / 5 ★</div>` : ''}
        <p class="address">${escapeHTML(business.address)}</p>
        <div class="business-details">
            <p>${escapeHTML(business.summary)}</p>
            <p><strong>Phone:</strong> ${escapeHTML(business.phone) || 'N/A'}</p>
            <p><strong>Website:</strong> ${business.website ? `<a href="${escapeHTML(business.website)}" target="_blank" rel="noopener noreferrer">${escapeHTML(business.website)}</a>` : 'N/A'}</p>
            <p><strong>Social:</strong> ${socialLinks}</p>
        </div>
    `;
    card.addEventListener('click', () => showDetailsModal(business));
    card.addEventListener('mouseenter', () => highlightMarker(business, true));
    card.addEventListener('mouseleave', () => highlightMarker(business, false));
    return card;
}

// --- Map Logic ---

function initMap() {
    if (map) return;
    map = L.map(mapContainer).setView([37.7749, -122.4194], 10); // Default to SF
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
    }).addTo(map);
    markersLayer = L.featureGroup().addTo(map);
}

function updateMapMarkers(businesses: any[]) {
    if (!map || !markersLayer) return;
    markersLayer.clearLayers();
    const markers: L.Marker[] = [];
    businesses.forEach(business => {
        if (business.lat && business.lng) {
            const marker = L.marker([business.lat, business.lng]);
            marker.bindPopup(`<b>${escapeHTML(business.name)}</b><br>${escapeHTML(business.address)}`);
            (marker as any).businessId = getBusinessId(business);
            markersLayer!.addLayer(marker);
            markers.push(marker);
        }
    });
    if (markers.length > 0) {
        map.fitBounds(L.featureGroup(markers).getBounds().pad(0.1));
    }
}

function highlightMarker(business: any, isHighlighted: boolean) {
    if (!markersLayer) return;
    const businessId = getBusinessId(business);
    markersLayer.eachLayer((layer: any) => {
        if (layer.businessId === businessId) {
            if (isHighlighted) {
                layer.openPopup();
            } else {
                layer.closePopup();
            }
        }
    });
}

// --- Filtering ---
function filterAndRenderResults() {
    const searchTerm = searchInput.value.toLowerCase();
    const instagramSearchTerm = instagramHandleInput.value.toLowerCase();
    const hasInstagramOnly = followsFilterCheckbox.checked;

    const filtered = currentBusinesses.filter(b => {
        const fullText = JSON.stringify(b).toLowerCase();
        const matchesSearch = searchTerm ? fullText.includes(searchTerm) : true;
        const matchesInstagram = instagramSearchTerm ? (b.socialMedia?.instagram || '').toLowerCase().includes(instagramSearchTerm) : true;
        const matchesHasInstagram = hasInstagramOnly ? !!b.socialMedia?.instagram : true;
        return matchesSearch && matchesInstagram && matchesHasInstagram;
    });
    renderResults(filtered);
}

// --- Export/Save ---

function exportToCsv() {
    const csvContent = convertToCsv(currentBusinesses);
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", "scout-ai-leads.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showStatus('CSV file downloaded.', 'success');
}

async function saveToSheets() {
    if (currentBusinesses.length === 0) {
        showStatus('No data to save.', 'error');
        return;
    }
    setButtonLoadingState(saveToSheetsButton, true);
    try {
        const sheetResponse = await gapi.client.sheets.spreadsheets.create({
            properties: { title: `Scout AI Leads - ${new Date().toLocaleString()}` }
        });
        const spreadsheetId = sheetResponse.result.spreadsheetId;
        const csv = convertToCsv(currentBusinesses);
        const data = csv.split('\n').map(row => row.split(','));

        await gapi.client.sheets.spreadsheets.values.update({
            spreadsheetId,
            range: 'A1',
            valueInputOption: 'USER_ENTERED',
            resource: { values: data }
        });
        showStatus(`Successfully saved to <a href="${sheetResponse.result.spreadsheetUrl}" target="_blank">Google Sheet</a>.`, 'success');
    } catch (error) {
        console.error('Save to Sheets Error:', error);
        showStatus('Failed to save to Google Sheets.', 'error');
    } finally {
        setButtonLoadingState(saveToSheetsButton, false);
    }
}

async function saveToDrive() {
    if (currentBusinesses.length === 0) {
        showStatus('No data to save.', 'error');
        return;
    }
    setButtonLoadingState(saveToDriveButton, true);
    const csvContent = convertToCsv(currentBusinesses);
    const fileName = `scout-ai-leads-${new Date().toISOString()}.csv`;
    const metadata = {
        name: fileName,
        mimeType: 'text/csv'
    };
    const boundary = '-------314159265358979323846';
    const delimiter = `\r\n--${boundary}\r\n`;
    const close_delim = `\r\n--${boundary}--`;

    const multipartRequestBody =
        delimiter +
        'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
        JSON.stringify(metadata) +
        delimiter +
        'Content-Type: text/csv\r\n\r\n' +
        csvContent +
        close_delim;

    try {
        const driveResponse = await gapi.client.request({
            path: '/upload/drive/v3/files',
            method: 'POST',
            params: { uploadType: 'multipart' },
            headers: {
                'Content-Type': `multipart/related; boundary="${boundary}"`
            },
            body: multipartRequestBody
        });
        showStatus(`CSV saved to Google Drive as "${fileName}".`, 'success');
    } catch (error) {
        console.error('Save to Drive Error:', error);
        showStatus('Failed to save CSV to Google Drive.', 'error');
    } finally {
        setButtonLoadingState(saveToDriveButton, false);
    }
}


// --- Modal ---
function showDetailsModal(business: any) {
    modalBusinessName.textContent = business.name;
    const socialLinks = business.socialMedia ? Object.entries(business.socialMedia).map(([platform, url]) => {
        if (!url) return '';
        return `<a href="${escapeHTML(url)}" target="_blank" rel="noopener noreferrer" class="social-link" aria-label="${platform}">${platform.charAt(0).toUpperCase() + platform.slice(1)}</a>`;
    }).join('') : 'N/A';

    modalBody.innerHTML = `
        <p><strong>Address:</strong> ${escapeHTML(business.address)}</p>
        ${business.rating ? `<p><strong>Rating:</strong> ${escapeHTML(business.rating)} / 5 ★</p>` : ''}
        <p>${escapeHTML(business.summary)}</p>
        <p><strong>Phone:</strong> ${escapeHTML(business.phone) || 'N/A'}</p>
        <p><strong>Website:</strong> ${business.website ? `<a href="${escapeHTML(business.website)}" target="_blank" rel="noopener noreferrer">${escapeHTML(business.website)}</a>` : 'N/A'}</p>
        <p><strong>Social Media:</strong> ${socialLinks}</p>
    `;
    detailsModal.classList.remove('hidden');
}

function hideDetailsModal() {
    detailsModal.classList.add('hidden');
}

// --- Event Listeners ---
aiSearchButton.addEventListener('click', handleAiSearch);
exportCsvButton.addEventListener('click', exportToCsv);
saveToSheetsButton.addEventListener('click', saveToSheets);
saveToDriveButton.addEventListener('click', saveToDrive);

[searchInput, instagramHandleInput].forEach(input => input.addEventListener('input', filterAndRenderResults));
followsFilterCheckbox.addEventListener('change', filterAndRenderResults);

modalCloseButton.addEventListener('click', hideDetailsModal);
detailsModal.addEventListener('click', (e) => {
    if (e.target === detailsModal) {
        hideDetailsModal();
    }
});
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !detailsModal.classList.contains('hidden')) {
        hideDetailsModal();
    }
});

googleClientIdInput.addEventListener('input', () => {
    const button = document.getElementById('authorize-button') as HTMLButtonElement;
    if (button) {
        button.disabled = !googleClientIdInput.value.trim();
    }
});


// --- App Initialization ---
function init() {
    // Load Google authentication scripts dynamically to prevent race conditions
    const gapiScript = document.createElement('script');
    gapiScript.src = 'https://apis.google.com/js/api.js';
    gapiScript.async = true;
    gapiScript.onload = () => gapi.load('client', initializeGapiClient);
    document.body.appendChild(gapiScript);

    const gisScript = document.createElement('script');
    gisScript.src = 'https://accounts.google.com/gsi/client';
    gisScript.async = true;
    gisScript.onload = () => {
        gisInited = true;
        maybeEnableAuth();
    };
    document.body.appendChild(gisScript);
}

init();