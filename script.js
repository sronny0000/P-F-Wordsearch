/* === CONFIGURATION CONSTANTS (Matching VBA) === */
const INPUT_RANGE = "A1:T15";
const ORIGINAL_SHEET_NAME = "OriginalData";
const ORIGINAL_SHEET_NAME_2 = "OriginalData2";
const COUNTER_CELL = "V2";

// =================================================================
// INTERNAL DATA SOURCES (20x15 Arrays - Replacing LocalStorage)
// =================================================================

// 1. Data Source for 'OriginalData' (A1:T15 - Matches the initial HTML table content)
const ORIGINAL_DATA_1_INTERNAL_RANGE = [
    ["Q", "V", "U", "C", "A", "L", "I", "L", "E", "N", "I", "N", "E", "T", "E", "E", "N", "B", "M", "A"],
    ["C", "R", "E", "T", "E", "H", "T", "R", "E", "A", "O", "U", "C", "S", "I", "F", "E", "V", "L", "E"],
    ["O", "R", "O", "U", "N", "D", "H", "N", "I", "M", "I", "H", "S", "T", "O", "L", "I", "A", "R", "O"],
    ["U", "H", "T", "R", "O", "N", "R", "O", "E", "O", "C", "E", "H", "S", "R", "A", "R", "D", "W", "O"],
    ["N", "O", "U", "B", "T", "D", "E", "L", "O", "E", "L", "T", "L", "N", "T", "R", "E", "T", "A", "W"],
    ["T", "E", "T", "A", "T", "S", "E", "O", "N", "Y", "N", "E", "O", "A", "O", "L", "I", "N", "E", "O"],
    ["P", "M", "U", "E", "H", "T", "T", "S", "E", "W", "O", "D", "A", "H", "R", "E", "E", "R", "E", "T"],
    ["E", "O", "N", "S", "K", "A", "E", "P", "T", "O", "O", "O", "I", "L", "P", "A", "S", "I", "E", "C"],
    ["L", "E", "S", "D", "B", "N", "O", "O", "R", "A", "O", "O", "E", "V", "L", "E", "W", "T", "S", "R"],
    ["E", "N", "T", "U", "F", "I", "V", "E", "T", "N", "T", "N", "C", "O", "I", "U", "J", "T", "S", "O"],
    ["V", "R", "N", "I", "E", "M", "A", "K", "L", "A", "W", "E", "S", "C", "O", "D", "T", "R", "I", "D"],
    ["E", "A", "D", "T", "T", "E", "Y", "O", "S", "S", "Q", "U", "A", "R", "E", "E", "E", "R", "H", "T"],
    ["N", "M", "T", "O", "P", "L", "E", "C", "E", "O", "H", "E", "T", "Q", "U", "A", "R", "R", "Y", "O"],
    ["N", "S", "I", "O", "S", "M", "I", "X", "Z", "F", "D", "Q", "B", "C", "C", "A", "S", "L", "I", "Y"],
    ["A", "P", "S", "O", "R", "E", "Z", "N", "G", "Q", "Z", "P", "S", "A", "P", "P", "H", "I", "R", "E"]
];

// 2. Data Source for 'OriginalData2' (Simulates a second sheet/range)
const ORIGINAL_DATA_2_INTERNAL_RANGE = [
    ["Q", "V", "U", "C", "A", "L", "I", "L", "E", "N", "I", "N", "E", "T", "E", "E", "N", "B", "M", "A"],
    ["C", "T", "H", "E", "R", "E", "T", "A", "R", "E", "O", "F", "I", "V", "E", "C", "L", "U", "E", "S"],
    ["O", "R", "O", "U", "N", "D", "H", "I", "N", "M", "T", "H", "I", "S", "O", "L", "I", "A", "R", "O"],
    ["U", "H", "T", "R", "O", "N", "R", "O", "E", "O", "W", "O", "R", "D", "S", "E", "A", "R", "C", "H"],
    ["N", "B", "U", "T", "D", "O", "E", "L", "N", "O", "T", "T", "E", "L", "L", "R", "E", "T", "A", "W"],
    ["T", "E", "T", "A", "T", "S", "E", "O", "A", "N", "Y", "O", "N", "E", "O", "L", "I", "N", "E", "O"],
    ["T", "H", "E", "M", "U", "P", "T", "S", "E", "W", "O", "D", "T", "H", "R", "E", "E", "A", "R", "E"],
    ["E", "O", "N", "S", "K", "A", "E", "P", "T", "O", "O", "O", "I", "S", "P", "E", "C", "I", "A", "L"],
    ["L", "B", "E", "N", "D", "S", "O", "O", "R", "A", "O", "O", "E", "V", "L", "E", "W", "T", "T", "U"],
    ["E", "R", "N", "S", "F", "I", "V", "E", "T", "N", "T", "N", "C", "O", "I", "J", "U", "S", "T", "O"],
    ["V", "R", "E", "M", "A", "I", "N", "K", "L", "A", "W", "E", "S", "C", "O", "D", "T", "R", "I", "D"],
    ["E", "S", "T", "E", "A", "D", "Y", "T", "O", "S", "Q", "U", "A", "R", "E", "E", "E", "R", "H", "T"],
    ["N", "C", "O", "M", "P", "L", "E", "T", "E", "O", "T", "H", "E", "Q", "U", "A", "R", "R", "Y", "O"],
    ["M", "I", "S", "S", "I", "O", "N", "X", "Z", "F", "D", "Q", "P", "L", "A", "Y", "B", "A", "S", "I"],
    ["A", "P", "S", "O", "R", "E", "Z", "N", "G", "Q", "Z", "P", "C", "C", "I", "P", "H", "E", "R", "S"]
];


// =================================================================
// VBA Function 1: Core Substitution Logic 
// =================================================================
function getSubstitution(inputLetter, direction) {
    const CYCLE1 = "-XEVALTF-";
    const CYCLE2 = "-JBG-";
    const CYCLE3 = "-QDW-";
    const CYCLE4 = "-ZK-";
    const CYCLE5 = "HRIO";
    const CYCLE6 = "MPS";
    const CYCLE7 = "CU";
    const CYCLE8 = "NY";
    const cycles = [CYCLE1, CYCLE2, CYCLE3, CYCLE4, CYCLE5, CYCLE6, CYCLE7, CYCLE8];

    let upperLetter = inputLetter.toUpperCase();
    let cycleStr = "";
    let pos = -1;

    for (let i = 0; i < cycles.length; i++) {
        const currentCycle = cycles[i];
        if (currentCycle.includes(upperLetter)) {
            cycleStr = currentCycle;
            pos = currentCycle.indexOf(upperLetter);
            break;
        }
    }
    if (cycleStr === "") return inputLetter;

    const cycleLen = cycleStr.length;
    const steps = 1;

    let shift = 0;
    if (direction.toUpperCase() === "RIGHT") {
        shift = steps;
    } else if (direction.toUpperCase() === "LEFT") {
        shift = -steps;
    }

    let nextPos = (pos + shift) % cycleLen;
    if (nextPos < 0) nextPos = nextPos + cycleLen;
    
    return cycleStr.charAt(nextPos);
}


// =================================================================
// VBA Subroutine 3: Core Execution Logic (applySubstitution)
// =================================================================
function applySubstitution(direction) {
    const grid = document.getElementById('data-grid');
    const cells = grid.querySelectorAll('td');
    
    // Update Counter Logic (VBA: PerformSubstitutionLeft/Right)
    const counterElement = document.getElementById('shift-counter');
    let currentCount = parseInt(counterElement.textContent) || 0;
    if (direction.toUpperCase() === "RIGHT") {
        currentCount -= 1; // Subtract 1 for Right 
    } else if (direction.toUpperCase() === "LEFT") {
        currentCount += 1; // Add 1 for Left
    }
    counterElement.textContent = currentCount;

    // Apply substitution to the grid
    cells.forEach(cell => {
        const startLetter = cell.textContent.trim();
        
        // Only process single letters
        if (startLetter.length === 1 && /[A-Za-z]/.test(startLetter)) {
            const currentLetter = getSubstitution(startLetter, direction);
            
            // Coloring Logic: Red if result is '-' (which happens in cycles 1-4)
            if (currentLetter === "-") {
                cell.style.backgroundColor = 'rgb(255, 0, 0)'; 
            } else {
                // Clear any previous color unless it's a Roman Numeral highlight (Blue)
                if (cell.style.backgroundColor !== 'rgb(0, 0, 255)') {
                   cell.style.backgroundColor = ''; 
                }
            }

            cell.textContent = currentLetter;
        }
    });
}


// =================================================================
// VBA Subroutine 5: RESET Data Functions (Uses internal constants)
// =================================================================

/**
 * Hides all image overlays.
 */
function hideAllOverlays() {
    document.getElementById('overlay-bitcoin').style.display = 'none';
    document.getElementById('overlay-oasis').style.display = 'none';
}


// Helper function to apply a 2D array of data to the HTML grid
function applyDataToGrid(data) {
    const grid = document.getElementById('data-grid');
    const cells = grid.querySelectorAll('td');
    
    // Flatten the 2D array into a 1D array of values
    const flatData = data.flat();

    // Copy data back from the data source to the grid
    cells.forEach((cell, index) => {
        // Ensure index is within bounds of stored data
        if (flatData[index]) {
            cell.textContent = flatData[index];
        } else {
             // Safety clear for unexpected index errors
            cell.textContent = ''; 
        }
        // Clear all color formatting on reset
        cell.style.backgroundColor = ''; 
    });

    // Reset the shift counter to 0
    const counterElement = document.getElementById('shift-counter');
    counterElement.textContent = 0; 
    counterElement.style.color = 'black'; // Reset color

    // HIDE OVERLAYS ON RESET
    hideAllOverlays();

    console.log(`The grid has been reset.`); 
}

function resetToOriginal() {
    // Reset uses the original data from the HTML table structure (ORIGINAL_DATA_1_INTERNAL_RANGE)
    applyDataToGrid(ORIGINAL_DATA_1_INTERNAL_RANGE); 
}

function resetToOriginal2() {
    // Reset uses the second internal data set (ORIGINAL_DATA_2_INTERNAL_RANGE)
    applyDataToGrid(ORIGINAL_DATA_2_INTERNAL_RANGE); 
}


// =================================================================
// VBA Subroutine 6: Roman Numeral Highlighting
// =================================================================
function isRomanNumeral(strValue) {
    let upperValue = strValue.toUpperCase();
    if (upperValue.length !== 1) return false;
    return "IVXLCDM".includes(upperValue); 
}

function highlightRomanNumerals() {
    const grid = document.getElementById('data-grid');
    const cells = grid.querySelectorAll('td');
    
    cells.forEach(cell => {
        let cellValue = cell.textContent.trim();
        
        // Check for Roman numeral and set color to Blue.
        if (isRomanNumeral(cellValue)) {
            cell.style.backgroundColor = 'rgb(0, 0, 255)'; // Blue
        } else {
            // No action taken if not a Roman Numeral, thus preserving existing color.
        }
    });
}


// =================================================================
// UPDATED: Image Overlay Logic
// =================================================================

/**
 * Toggles the display of a specific image overlay without affecting others.
 * @param {string} name - The name of the overlay to toggle ('bitcoin' or 'oasis').
 */
function toggleOverlay(name) {
    let overlayElement;
    if (name === 'bitcoin') {
        overlayElement = document.getElementById('overlay-bitcoin');
    } else if (name === 'oasis') {
        overlayElement = document.getElementById('overlay-oasis');
    }

    if (overlayElement) {
        // Toggle visibility: 'block' (visible) or 'none' (hidden)
        if (overlayElement.style.display === 'block') {
            overlayElement.style.display = 'none';
        } else {
            overlayElement.style.display = 'block';
        }
    }
}