/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */
import * as L from 'leaflet';
import { GoogleGenAI, Type } from '@google/genai';

// --- DOM Elements ---
const authorizeButton = document.getElementById('authorize-button') as HTMLButtonElement;
const signoutButton = document.getElementById('signout-button') as HTMLButtonElement;
const sheetUrlInput = document.getElementById('sheet-url-input') as HTMLInputElement;
const loadSheetButton = document.getElementById('load-sheet-button') as HTMLButtonElement;
const cityInput = document.getElementById('city-input') as HTMLInputElement;
const stateInput = document.getElementById('state-input') as HTMLInputElement;
const aiSearchButton = document.getElementById('ai-search-button') as HTMLButtonElement;
const exportButton = document.getElementById('export-button') as HTMLButtonElement;
const resultsContainer = document.getElementById('results-container') as HTMLDivElement;
const loader = document.getElementById('loader') as HTMLDivElement;
const viewFavoritesButton = document.getElementById('view-favorites-button') as HTMLButtonElement;
const favoritesContainer = document.getElementById('favorites-container') as HTMLDivElement;
const favoritesList = document.getElementById('favorites-list') as HTMLDivElement;
const app = document.getElementById('app') as HTMLDivElement;
const mapContainer = document.getElementById('map') as HTMLDivElement;
const authContainer = document.getElementById('auth-container') as HTMLDivElement;
const sheetContainer = document.getElementById('sheet-container') as HTMLDivElement;
const searchInput = document.getElementById('search-input') as HTMLInputElement;
const searchContainer = document.getElementById('search-container') as HTMLDivElement;
const sortContainer = document.getElementById('sort-container') as HTMLDivElement;
const sortBySelect = document.getElementById('sort-by-select') as HTMLSelectElement;
const sortDirectionButton = document.getElementById('sort-direction-button') as HTMLButtonElement;

// --- Google API Config ---
const CLIENT_ID = 'YOUR_CLIENT_ID.apps.googleusercontent.com'; // IMPORTANT: Replace with your actual Client ID
const API_KEY = process.env.API_KEY;
const DISCOVERY_DOCS = ["https://sheets.googleapis.com/$discovery/rest?version=v4"];
const SCOPES = "https://www.googleapis.com/auth/spreadsheets";

let gapi: any;
let google: any;
let tokenClient: any;
const ai = new GoogleGenAI({ apiKey: API_KEY });


// --- App State ---
const LOCAL_STORAGE_KEY = 'lastSheetUrl';
let map: L.Map | null = null;
let markersLayer: L.FeatureGroup | null = null;
let currentSalons: any[] = [];
let headers: string[] = [];
let currentSpreadsheetId: string | null = null;
let dataPollingInterval: number | null = null;
let lastDataState: string = '';
let isEditing = false; // Flag to prevent polling during edits
let currentSortKey: string = 'none';
let currentSortDirection: 'asc' | 'desc' = 'asc';

// --- Error Handling ---

/**
 * Parses a GAPI error object and returns a user-friendly string.
 * @param err The error object from a GAPI client request.
 * @returns A user-friendly error message.
 */
function parseGapiError(err: any): string {
    if (err && err.result && err.result.error) {
        const error = err.result.error;
        switch (error.status) {
            case 'NOT_FOUND':
                return 'The Google Sheet could not be found. Please check the URL or ID and make sure it\'s correct.';
            case 'PERMISSION_DENIED':
                return 'You do not have permission to access this Google Sheet. Please check the sheet\'s sharing settings and ensure you are signed in with the correct Google account.';
            case 'RESOURCE_EXHAUSTED':
                return 'The app has made too many requests to Google Sheets. Please wait a moment and try again later.';
            default:
                return error.message || 'An unknown error occurred while communicating with Google Sheets.';
        }
    }
    return 'An unexpected error occurred. Please check the console for more details.';
}


// --- Main Functions ---

/**
 * Called after the Google API script has loaded.
 */
(window as any).gapiLoaded = () => {
    gapi = (window as any).gapi;
    gapi.load('client', initializeGapiClient);
};

/**
 * Initializes the GAPI client and sets up the token client for auth.
 */
async function initializeGapiClient() {
    try {
        await gapi.client.init({
            apiKey: API_KEY,
            discoveryDocs: DISCOVERY_DOCS,
        });

        google = (window as any).google;
        if (!google || !google.accounts) {
            console.error('Google Identity Services library not loaded.');
            authorizeButton.textContent = 'Auth failed to load';
            authorizeButton.disabled = true;
            return;
        }

        tokenClient = google.accounts.oauth2.initTokenClient({
            client_id: CLIENT_ID,
            scope: SCOPES,
            callback: '', // defined dynamically
        });

        authorizeButton.disabled = false;
        authorizeButton.textContent = 'Connect Google Sheets';
    } catch (err) {
        console.error('Error during GAPI client initialization:', err);
        authorizeButton.textContent = 'Auth initialization failed';
        authorizeButton.disabled = true;
    }
}

/**
 * Handles the authorization process.
 */
function handleAuthClick() {
    if (!tokenClient) {
        alert('Authentication client is not ready. Please wait a moment and try again.');
        console.error('handleAuthClick called before tokenClient was initialized.');
        return;
    }

    tokenClient.callback = async (resp: any) => {
        if (resp.error !== undefined) {
            throw (resp);
        }
        updateSigninStatus(true);
    };

    if (gapi.client.getToken() === null) {
        tokenClient.requestAccessToken({ prompt: 'consent' });
    } else {
        tokenClient.requestAccessToken({ prompt: '' });
    }
}

/**
 * Handles user sign-out.
 */
function handleSignoutClick() {
    // Fix: Corrected typo from 'gpi' to 'gapi'.
    const token = gapi.client.getToken();
    if (token !== null) {
        google.accounts.oauth2.revoke(token.access_token, () => {
            gapi.client.setToken('');
            updateSigninStatus(false);
        });
    }
}

/**
 * Updates the UI based on the user's sign-in status.
 */
function updateSigninStatus(isSignedIn: boolean) {
    if (isSignedIn) {
        authContainer.classList.add('hidden');
        signoutButton.classList.remove('hidden');
        sheetContainer.classList.remove('hidden');
    } else {
        authContainer.classList.remove('hidden');
        signoutButton.classList.add('hidden');
        sheetContainer.classList.add('hidden');
        resultsContainer.innerHTML = '<p>Please connect your Google Account to load a sheet.</p>';
        mapContainer.classList.add('hidden');
        if (dataPollingInterval) {
            clearInterval(dataPollingInterval);
            dataPollingInterval = null;
        }
        currentSpreadsheetId = null;
        currentSalons = [];
        sortContainer.classList.add('hidden');
    }
}

/**
 * Extracts the spreadsheet ID from a Google Sheet URL.
 */
function getSpreadsheetIdFromUrl(url: string): string | null {
    const match = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/.exec(url);
    return match ? match[1] : url; // Assume it's an ID if no match
}

/**
 * Main function to load data from the specified Google Sheet.
 */
async function loadData() {
    if (isEditing) {
        console.log('Polling skipped: user is editing.');
        return;
    }
    if (!currentSpreadsheetId) {
        return; // Silently exit if no sheet is loaded
    }
    setLoading(true);
    try {
        const range = 'Sheet1!A:Z'; // Read all data
        const response = await gapi.client.sheets.spreadsheets.values.get({
            spreadsheetId: currentSpreadsheetId,
            range: range,
        });

        const data = response.result.values;
        if (!data || data.length < 2) {
            displayError("No data found in the sheet, or sheet is missing a header row.");
            setLoading(false);
            return;
        }

        const currentState = JSON.stringify(data);
        if (currentState === lastDataState) {
            console.log('Data unchanged. Skipping re-render.');
            setLoading(false);
            return;
        }
        lastDataState = currentState;

        const newHeaders = data[0];
        if (JSON.stringify(headers) !== JSON.stringify(newHeaders)) {
            headers = newHeaders;
            updateSortOptions(headers);
        }

        const rows = data.slice(1);

        currentSalons = rows.map((row: any[], index: number) => {
            const salon: any = { rowIndex: index + 2 }; // +2 for header and 1-based indexing
            headers.forEach((header, i) => {
                salon[header.toLowerCase()] = row[i] || '';
            });
            return salon;
        });

        applyFilterAndRender();

    } catch (err: any) {
        console.error('Error loading sheet data:', err);
        const friendlyMessage = parseGapiError(err);
        displayError(friendlyMessage);

        const status = err?.result?.error?.status;
        if (status === 'NOT_FOUND' || status === 'PERMISSION_DENIED') {
            if (dataPollingInterval) {
                clearInterval(dataPollingInterval);
                dataPollingInterval = null;
                console.warn('Polling stopped due to a fatal error.');
            }
        }
    } finally {
        setLoading(false);
    }
}

/**
 * Updates a specific range in the Google Sheet.
 */
async function updateSheetData(range: string, values: any[][]) {
    await gapi.client.sheets.spreadsheets.values.update({
        spreadsheetId: currentSpreadsheetId,
        range: range,
        valueInputOption: 'USER_ENTERED',
        resource: {
            values: values,
        },
    });
}

const getSalonId = (salon: any): string => btoa(salon.name + salon.address);

const isFavorited = (salon: any): boolean => {
    const favValue = salon.favorite ? salon.favorite.toString().toUpperCase() : 'FALSE';
    return favValue === 'TRUE';
};

const toggleFavorite = (salon: any) => {
    // This function only works for salons loaded from a sheet with a rowIndex
    if (salon.rowIndex === undefined) {
        alert("You can only favorite items that are loaded from a Google Sheet.");
        return;
    }
    const originalFavoriteState = salon.favorite;
    const isNowFavorited = !isFavorited(salon);
    salon.favorite = isNowFavorited ? 'TRUE' : 'FALSE';

    updateFavoriteButtons(salon);

    const favoriteColIndex = headers.findIndex(h => h.toLowerCase() === 'favorite');
    if (favoriteColIndex === -1) {
        alert('Could not find a "favorite" column in your sheet.');
        salon.favorite = originalFavoriteState; // revert optimistic change
        updateFavoriteButtons(salon);
        return;
    }
    const colLetter = String.fromCharCode('A'.charCodeAt(0) + favoriteColIndex);
    const range = `Sheet1!${colLetter}${salon.rowIndex}`;

    updateSheetData(range, [[salon.favorite]]).then(() => {
        const index = currentSalons.findIndex(s => s.rowIndex === salon.rowIndex);
        if (index !== -1) {
            currentSalons[index].favorite = salon.favorite;
            const card = document.getElementById(`card-${getSalonId(salon)}`);
            if (card) {
                card.dataset.salonInfo = JSON.stringify(currentSalons[index]);
            }
        }
        if (!favoritesContainer.classList.contains('hidden')) {
            displayFavorites();
        }
    }).catch((err) => {
        console.error('Favorite update failed:', err);
        const friendlyMessage = parseGapiError(err);
        alert(`Failed to update favorite status: ${friendlyMessage}`);
        salon.favorite = originalFavoriteState;
        updateFavoriteButtons(salon);
    });
};

const updateFavoriteButtons = (salon: any) => {
    const salonId = getSalonId(salon);
    const isNowFavorited = isFavorited(salon);
    const buttons = document.querySelectorAll(`.favorite-button[data-salon-id="${salonId}"]`);

    buttons.forEach((button: HTMLButtonElement) => {
        button.textContent = isNowFavorited ? 'Unfavorite' : 'Favorite';
        button.classList.toggle('favorited', isNowFavorited);
    });
};

function setLoading(isLoading: boolean) {
    loader.classList.toggle('hidden', !isLoading);
}

function displayError(message: string) {
    resultsContainer.innerHTML = `<div class="error">${message}</div>`;
}

// --- Card and Display Functions ---

const socialIcons: { [key: string]: string } = {
    instagram: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M12 2.163c3.204 0 3.584.012 4.85.07 3.252.148 4.771 1.691 4.919 4.919.058 1.265.069 1.645.069 4.85s-.011 3.584-.069 4.85c-.149 3.225-1.664 4.771-4.919 4.919-1.266.058-1.644.07-4.85.07s-3.584-.012-4.85-.07c-3.252-.148-4.771-1.691-4.919-4.919-.058-1.265-.069-1.645-.069-4.85s.011-3.584.069-4.85c.149-3.225 1.664-4.771 4.919-4.919C8.416 2.175 8.796 2.163 12 2.163zm0 1.441c-3.171 0-3.53.011-4.76.069-2.776.127-4.001 1.348-4.129 4.129-.058 1.23-.069 1.589-.069 4.76s.011 3.53.069 4.76c.127 2.776 1.353 4.001 4.129 4.129 1.23.058 1.589.069 4.76.069s3.53-.011 4.76-.069c2.776-.127 4.001-1.348 4.129-4.129.058-1.23.069-1.589.069-4.76s-.011-3.53-.069-4.76c-.127-2.776-1.353-4.001-4.129-4.129-1.23-.058-1.589-.069-4.76-.069zm0 4.088a4.297 4.297 0 100 8.594 4.297 4.297 0 000-8.594zm0 7.152a2.855 2.855 0 110-5.71 2.855 2.855 0 010 5.71zm4.398-7.882a1.033 1.033 0 100 2.066 1.033 1.033 0 000-2.066z"/></svg>`,
    facebook: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M22 12c0-5.523-4.477-10-10-10S2 6.477 2 12c0 4.991 3.657 9.128 8.438 9.879V14.89H8.078v-2.906h2.36V9.627c0-2.343 1.402-3.625 3.518-3.625 1.002 0 1.864.074 2.115.108v2.589h-1.52c-1.135 0-1.354.538-1.354 1.332v1.75h2.87l-.372 2.906h-2.498v7.005C18.343 21.128 22 16.991 22 12z"/></svg>`,
    tiktok: `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M12.525 2.022c1.895 0 3.492.622 4.745 1.875 1.253 1.253 1.875 2.85 1.875 4.745v.078c0 .281.01.552.03.812H16.89a6.602 6.602 0 0 1-.094-.89c0-1.406-.547-2.64-1.64-3.735S12.568 3.428 11.16 3.428v3.428h-3.47a3.428 3.428 0 1 1 0-6.857h4.836zm-3.47 11.75v-2.062H5.625v2.062a4.812 4.812 0 0 1-4.812 4.813A4.812 4.812 0 0 1 0 18.583a4.812 4.812 0 0 1 4.813-4.812c.078 0 .156 0 .234.01V2.023h3.429v11.75z"/></svg>`,
};

function renderSocials(salon: any): string {
    const collectedSocials: { [key: string]: string } = { ...salon.social_media };
    
    // Also check for individual columns from sheets
    for (const key in salon) {
        if (socialIcons[key] && !collectedSocials[key]) {
            collectedSocials[key] = salon[key];
        }
    }

    if (Object.keys(collectedSocials).length === 0) return '';

    const links = Object.entries(collectedSocials)
        .filter(([_, url]) => url && url.startsWith('http'))
        .map(([platform, url]) => {
            if (socialIcons[platform]) {
                return `<a href="${escapeHTML(url)}" class="social-link" target="_blank" rel="noopener noreferrer" title="${platform.charAt(0).toUpperCase() + platform.slice(1)}">${socialIcons[platform]}</a>`;
            }
            return '';
        }).join('');

    return links ? `<div class="social-links">${links}</div>` : '';
}

function renderSources(salon: any): string {
    if (!salon.sources || salon.sources.length === 0) return '';

    const links = salon.sources
        .filter((url: string) => url && url.startsWith('http'))
        .map((url: string) => {
            let domain = 'Source';
            try {
                domain = new URL(url).hostname.replace('www.', '');
            } catch (e) { /* ignore invalid URLs */ }
            return `<a href="${escapeHTML(url)}" class="source-link" target="_blank" rel="noopener noreferrer">${escapeHTML(domain)}</a>`;
        }).join(', ');

    return links ? `<div class="sources-container"><strong>Sources:</strong> ${links}</div>` : '';
}

function createSalonCards(salons: any[]): string {
    return salons.map(salon => {
        const salonId = getSalonId(salon);
        const isFav = isFavorited(salon);
        const salonInfo = escapeHTML(JSON.stringify(salon));
        const canEdit = salon.rowIndex !== undefined; // Can only edit if it has a row index from a sheet

        // Create a dedicated section for Hours and Website
        const detailsHTML = [
            salon.hours ? `<p class="hours" data-field="hours" contenteditable="false"><strong>Hours:</strong> ${escapeHTML(salon.hours)}</p>` : '',
            salon.website ? `<p class="website"><strong>Website:</strong> <a href="${escapeHTML(salon.website)}" target="_blank" rel="noopener noreferrer">${escapeHTML(salon.website)}</a></p>` : ''
        ].filter(Boolean).join('');

        return `
        <div class="salon-card" id="card-${salonId}" role="article" data-salon-info="${salonInfo}">
            <h3 id="salon-name-${salonId}" data-field="name" contenteditable="false">${escapeHTML(salon.name)}</h3>
            ${salon.rating ? `<div class="rating" data-field="rating" contenteditable="false">⭐ ${escapeHTML(salon.rating)}</div>` : ''}
            <p class="address" data-field="address" contenteditable="false">${escapeHTML(salon.address)}</p>
            <p data-field="description" contenteditable="false">${escapeHTML(salon.description)}</p>
            
            ${detailsHTML ? `<div class="salon-details">${detailsHTML}</div>` : ''}
            
            ${salon.phone ? `<div class="contact-info"><a href="tel:${escapeHTML(salon.phone)}" class="salon-action-button">Call: ${escapeHTML(salon.phone)}</a></div>` : ''}
            ${renderSources(salon)}
            <div class="salon-card-actions">
                ${renderSocials(salon)}
                <button class="favorite-button ${isFav ? 'favorited' : ''}" data-salon-id="${salonId}" ${canEdit ? '' : 'disabled'}>${isFav ? 'Unfavorite' : 'Favorite'}</button>
                ${canEdit ? `
                    <button class="edit-button salon-action-button">Edit</button>
                    <button class="save-button salon-action-button hidden">Save</button>
                ` : ''}
            </div>
        </div>
    `}).join('');
}


function displayResults(salons: any[]) {
    if (!salons || salons.length === 0) {
        const query = searchInput.value.trim();
        if (query) {
             resultsContainer.innerHTML = `<p>No salons found matching "${escapeHTML(query)}".</p>`;
        } else {
            resultsContainer.innerHTML = '<p>No salons found.</p>';
        }
        updateMap([]); // Clear the map
        return;
    }
    resultsContainer.innerHTML = createSalonCards(salons);
    updateMap(salons);
}

function displayFavorites() {
    mapContainer.classList.add('hidden');
    resultsContainer.innerHTML = '';
    const favoriteSalons = currentSalons.filter(isFavorited);

    if (favoriteSalons.length === 0) {
        favoritesList.innerHTML = '<p>You haven\'t saved any favorite salons yet.</p>';
        return;
    }

    favoritesList.innerHTML = createSalonCards(favoriteSalons);
}

/**
 * Sorts an array of salon objects based on a key and direction.
 */
function sortSalons(salons: any[], key: string, direction: 'asc' | 'desc'): any[] {
    const sorted = [...salons];
    const dir = direction === 'asc' ? 1 : -1;

    sorted.sort((a, b) => {
        let valA = a[key] ?? '';
        let valB = b[key] ?? '';

        // Handle boolean-like strings 'TRUE'/'FALSE'
        if (typeof valA === 'string' && (valA.toUpperCase() === 'TRUE' || valA.toUpperCase() === 'FALSE')) {
            valA = valA.toUpperCase() === 'TRUE';
        }
        if (typeof valB === 'string' && (valB.toUpperCase() === 'TRUE' || valB.toUpperCase() === 'FALSE')) {
            valB = valB.toUpperCase() === 'TRUE';
        }

        // Attempt numeric comparison
        const numA = parseFloat(valA);
        const numB = parseFloat(valB);

        // Check if they are valid numbers and weren't just parsed from a string like "123 Main St"
        if (!isNaN(numA) && !isNaN(numB) && valA.toString() === numA.toString() && valB.toString() === numB.toString()) {
            if (numA < numB) return -1 * dir;
            if (numA > numB) return 1 * dir;
            return 0;
        } else {
            // Fallback to string comparison
            return valA.toString().localeCompare(valB.toString(), undefined, { numeric: true, sensitivity: 'base' }) * dir;
        }
    });
    return sorted;
}


function applyFilterAndRender() {
    if (favoritesContainer.classList.contains('hidden')) {
        const query = searchInput.value.trim().toLowerCase();
        let salonsToRender = currentSalons;

        if (query) {
            salonsToRender = currentSalons.filter(salon => {
                const name = (salon.name || '').toLowerCase();
                const address = (salon.address || '').toLowerCase();
                const description = (salon.description || '').toLowerCase();
                return name.includes(query) || address.includes(query) || description.includes(query);
            });
        }
        
        if (currentSortKey !== 'none') {
            salonsToRender = sortSalons(salonsToRender, currentSortKey, currentSortDirection);
        }

        displayResults(salonsToRender);
    }
}

function initMap(lat: number, lon: number) {
    if (map) {
        map.setView([lat, lon], 13);
        return;
    };
    map = L.map(mapContainer).setView([lat, lon], 13);
    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
    }).addTo(map);
    markersLayer = L.featureGroup().addTo(map);
}

function updateMap(salons: any[]) {
    const validSalons = salons.filter(s => s.latitude && s.longitude);

    if (validSalons.length === 0) {
        mapContainer.classList.add('hidden');
        return;
    }

    mapContainer.classList.remove('hidden');

    if (!map) {
        initMap(parseFloat(validSalons[0].latitude), parseFloat(validSalons[0].longitude));
    }

    setTimeout(() => map?.invalidateSize(), 100);

    markersLayer!.clearLayers();

    validSalons.forEach(salon => {
        const salonId = getSalonId(salon);
        const marker = L.marker([salon.latitude, salon.longitude], { salonId: salonId } as any);
        marker.bindPopup(`<b>${escapeHTML(salon.name)}</b><br>${escapeHTML(salon.address)}`);

        marker.on('click', () => {
            document.querySelectorAll('.salon-card.highlighted').forEach(el => el.classList.remove('highlighted'));
            const card = document.getElementById(`card-${salonId}`);
            if (card) {
                card.classList.add('highlighted');
                card.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
        });
        markersLayer!.addLayer(marker);
    });

    if (markersLayer!.getLayers().length > 0) {
        map!.