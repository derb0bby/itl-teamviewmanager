/**
 * Constants
 */
const Ansicht_CONFIGS_PROPERTY = "AnsichtViewConfigurations";
const DEFAULT_CONFIG = {
  "Ansicht 1": [],
}; // Default name for a new Ansicht
const EXCEL_LIMITS = {
  MAX_COLUMN_LENGTH: 3, // Maximum characters in column reference (AAA)
  MAX_COLUMN_NUMBER: 16384, // Excel's maximum column (XFD)
  LAST_COLUMN_NAME: "XFD", // Excel's last column name
};

// Modal Config
const isOpenClass = "modal-is-open";
const openingClass = "modal-is-opening";
const closingClass = "modal-is-closing";
const scrollbarWidthCssVar = "--pico-scrollbar-width";
const animationDuration = 400; // ms
let visibleModal = null;

/**
 * Converts Excel column letter to number (e.g., 'A' -> 1, 'AA' -> 27)
 * @param {string} column - Column letter (e.g., 'A', 'BC', 'XFD')
 * @returns {number} Column number
 */
function convertColumnLetterToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result *= 26;
    result += column.charCodeAt(i) - "A".charCodeAt(0) + 1;
  }
  return result;
}

/**
 * Validates a single column reference
 * @param {string} column - Single column reference to validate
 * @returns {Object} Validation result and error message
 */
function validateSingleColumn(column) {
  // Check column length
  if (column.length > EXCEL_LIMITS.MAX_COLUMN_LENGTH) {
    return {
      isValid: false,
      message: `Ungültige Spaltenreferenz: ${column}. Spalten dürfen maximal ${EXCEL_LIMITS.MAX_COLUMN_LENGTH} Buchstaben lang sein.`,
    };
  }

  // Check if column is within Excel's limits
  const columnNumber = convertColumnLetterToNumber(column);
  if (columnNumber > EXCEL_LIMITS.MAX_COLUMN_NUMBER) {
    return {
      isValid: false,
      message: `Spalte ${column} liegt außerhalb des gültigen Excel-Bereichs (max. ${EXCEL_LIMITS.LAST_COLUMN_NAME}).`,
    };
  }

  return {
    isValid: true,
  };
}

/**
 * Sanitizes column input by removing whitespace and converting to uppercase
 * @param {string} input - The column input string to sanitize
 * @returns {string} Sanitized input string
 */
function sanitizeColumnInput(input) {
  return input
    .replace(/\s+/g, "") // Remove all whitespace
    .toUpperCase() // Convert to uppercase
    .replace(/[^A-Z,]/g, "") // Remove any characters that aren't letters or commas
    .replace(/,+/g, ","); // Replace multiple consecutive commas with a single comma
  // .replace(/^,|,$/g, ""); // Remove leading and trailing commas
}

/**
 * Validates column input format
 * @param {string} input - The column input string to validate
 * @returns {Object} Object containing validation result and error message
 */
function validateColumnInput(input) {
  if (!input || input.trim() === "") {
    return {
      isValid: false,
      message: "Bitte geben Sie mindestens eine Spalte an.",
    };
  }

  // Remove any whitespace and convert to uppercase
  const sanitizedInput = sanitizeColumnInput(input);

  // Basic format check (letters and commas)
  const validFormat = /^[A-Z]+(,[A-Z]+)*$/.test(sanitizedInput);
  if (!validFormat) {
    return {
      isValid: false,
      message: "Ungültiges Format. Bitte verwenden Sie nur Buchstaben und Kommas (z.B., A,B,C).",
    };
  }

  // Split into individual columns
  const columns = sanitizedInput.split(",");

  // Validate each column reference
  for (const column of columns) {
    // Check for empty column references
    if (column.length === 0) {
      return {
        isValid: false,
        message: "Leere Spaltenreferenz gefunden. Bitte überprüfen Sie die Kommas (z.B., A,,B,C).",
      };
    }

    // Validate column format and limits
    const columnValidation = validateSingleColumn(column);
    if (!columnValidation.isValid) {
      return columnValidation;
    }
  }

  return {
    isValid: true,
    sanitizedValue: sanitizedInput,
    columns: columns,
  };
}

// Initialize Office JS
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Set up event handlers for existing buttons

    // document.getElementById("btnAnsichtLager").onclick = () => applyAnsichtView("Ansicht_Lager");
    // document.getElementById("btnAnsichtSM").onclick = () => applyAnsichtView("Ansicht_SM");
    // document.getElementById("btnAnsichtTechnik").onclick = () => applyAnsichtView("Ansicht_Technik");
    document.getElementById("btnConfigureView").onclick = showConfigurationDialog;
    document.getElementById("btnCloseDialog").onclick = hideConfigurationDialog;
    document.getElementById("btnCloseAnsichtManagementDialog").onclick = hideAnsichtManagementDialog;
    document.getElementById("btnSaveConfig").onclick = saveConfiguration;
    document.getElementById("btnAddSheet").onclick = addSheetConfiguration;
    document.getElementById("AnsichtSelect").onchange = loadAnsichtConfiguration;

    document.getElementById("btnAddAnsicht").onclick = addAnsicht;
    document.getElementById("btnRenameAnsicht").onclick = renameAnsicht;
    document.getElementById("btnDeleteAnsicht").onclick = deleteAnsicht;
    document.getElementById("btnCloseModal").onclick = toggleModal;
    document.getElementById("btnConfirmAddAnsicht").onclick = confirmAddAnsicht;
    document.getElementById("btnConfirmRenameAnsicht").onclick = confirmRenameAnsicht;

    updateAnsichtSelectOptions(); // Load existing configuration and create Ansicht buttons
  }
});

/**
 * Shows the configuration dialog and loads current settings
 */
async function showConfigurationDialog() {
  const dialog = document.getElementById("configDialog");
  dialog.style.display = "block";

  // Load current Ansicht's configuration
  const AnsichtSelect = document.getElementById("AnsichtSelect");
  await loadAnsichtConfiguration(AnsichtSelect.value);
}

/**
 * Hides the configuration dialog
 */
function hideConfigurationDialog() {
  document.getElementById("configDialog").style.display = "none";
  document.getElementById("AnsichtManagement").style.display = "none";
}

function hideAnsichtManagementDialog() {
  document.getElementById("AnsichtManagement").style.display = "none";
  document.getElementById("AnsichtInput").value = "";
}

/**
 * Loads configurations from workbook custom properties
 * @returns {Promise<Object>} The Ansicht configurations
 */

/**
 * Loads configurations from the Office.context.document.settings property bag.
 * Returns the stored configuration object or default values if none exist.
 */
async function loadSettingsFromStorage() {
  try {
    const storedConfig = Office.context.document.settings.get(Ansicht_CONFIGS_PROPERTY);

    if (!storedConfig) {
      // Return default empty configuration if no configuration exists
      return DEFAULT_CONFIG;
    }

    return JSON.parse(storedConfig); // Parse and return the stored configuration
  } catch (error) {
    console.error("Error loading configurations:", error);
    showError("Error loading configurations");
    return DEFAULT_CONFIG;
  }
}

/**
 * Saves configurations to the Office.context.document.settings property bag.
 * Checks for size limit of 5MB before saving.
 * @param {Object} configs - The configurations to save.
 */
async function saveSettingsToStorage(configs) {
  try {
    const jsonString = JSON.stringify(configs);

    // Check size limit (5 MB = 5 * 1024 * 1024 bytes)
    const maxBytes = 5 * 1024 * 1024;
    const jsonSize = new Blob([jsonString]).size;

    if (jsonSize > maxBytes) {
      const errorMessage = `Die Größe der Konfiguration überschreitet das mögliche Limit. Bitte reduzieren Sie die Anzahl an Ansichten und versuchen Sie es erneut.`;
      console.error(errorMessage);
      showError(errorMessage); // Display an error message to the user
      return;
    }

    // Save the JSON string as a single setting
    Office.context.document.settings.set(Ansicht_CONFIGS_PROPERTY, jsonString);

    // Persist the changes to the document
    Office.context.document.settings.saveAsync(function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Error saving configurations:", asyncResult.error.message);
        showError("Error saving configurations");
      } else {
        console.log("Configurations saved successfully.");
        hideConfigurationDialog(); // Close dialog after saving
      }
    });
  } catch (error) {
    console.error("Error saving configurations:", error);
    showError("Error saving configurations");
    throw error;
  }
}

/**
 * Loads configuration for a specific Ansicht
 */
async function loadAnsichtConfiguration() {
  const AnsichtKey = document.getElementById("AnsichtSelect").value;
  const configs = await loadSettingsFromStorage();
  const AnsichtConfig = configs[AnsichtKey] || [];

  // Get available sheets
  const sheets = await getWorksheetNames();

  // Clear existing configuration UI
  const sheetConfig = document.getElementById("sheetConfig");
  sheetConfig.innerHTML = "";

  // Create configuration UI for each sheet
  AnsichtConfig.forEach((config, index) => {
    addSheetConfigurationUI(sheetConfig, sheets, config, index);
  });
}

/**
 * Adds a new sheet configuration UI element
 */
async function addSheetConfiguration() {
  const sheets = await getWorksheetNames();
  const sheetConfig = document.getElementById("sheetConfig");
  addSheetConfigurationUI(sheetConfig, sheets, null, sheetConfig.children.length);
}

/**
 * Creates UI elements for sheet configuration with input validation
 * @param {HTMLElement} container - Container element
 * @param {string[]} sheets - Available worksheet names
 * @param {Object} config - Existing configuration
 * @param {number} index - Configuration index
 */
function addSheetConfigurationUI(container, sheets, config, index) {
  const div = document.createElement("div");
  div.className = "sheet-config";

  const columnsValue = config ? config.visibleColumns.join(",") : "";
  const sanitizedColumnsValue = config ? sanitizeColumnInput(columnsValue) : "";

  div.innerHTML = `
    <select class="sheet-select">
      ${sheets
        .map(
          (sheet) =>
            `<option value="${sheet}" ${config && config.sheetName === sheet ? "selected" : ""}>${sheet}</option>`
        )
        .join("")}
    </select>
    <input type="text" id="columnSpecification" class="columns-input" placeholder="Sichtbare Spalten (z.B., A,C,E)" 
           value="${sanitizedColumnsValue}">
    <div class="input-error" style="display: none; color: red; font-size: 0.8em;"></div>
    <button class="button-remove" onclick="removeSheetConfig(${index})">Löschen</button>
  `;

  // Add input validation event listener
  const columnsInput = div.querySelector(".columns-input");
  const errorDiv = div.querySelector(".input-error");

  columnsInput.addEventListener("input", (event) => {
    const validation = validateColumnInput(event.target.value);

    if (!validation.isValid) {
      errorDiv.textContent = validation.message;
      errorDiv.style.display = "block";
      columnsInput.classList.add("input-error");
    } else {
      errorDiv.style.display = "none";
      columnsInput.classList.remove("input-error");
      columnsInput.value = validation.sanitizedValue;
    }
  });

  container.appendChild(div);
}

/**
 * Removes a sheet configuration
 * @param {number} index - Configuration index to remove
 */
function removeSheetConfig(index) {
  const configs = document.getElementsByClassName("sheet-config");
  if (configs[index]) {
    configs[index].remove();
  }
}

/**
 * Gets all worksheet names from the current workbook
 * @returns {Promise<string[]>} Array of worksheet names
 */
async function getWorksheetNames() {
  try {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      return sheets.items.map((sheet) => sheet.name);
    });
  } catch (error) {
    console.error("Error getting worksheet names:", error);
    return [];
  }
}

/**
 * Saves the current configuration
 */
async function saveConfiguration() {
  try {
    const AnsichtKey = document.getElementById("AnsichtSelect").value;
    const sheetConfigs = document.getElementsByClassName("sheet-config");
    const configs = await loadSettingsFromStorage();

    // Validate all inputs before saving
    let hasErrors = false;
    const newConfigs = Array.from(sheetConfigs).map((config) => {
      const sheetName = config.querySelector(".sheet-select").value;
      const columnsInput = config.querySelector(".columns-input").value;

      const validation = validateColumnInput(columnsInput);
      if (!validation.isValid) {
        hasErrors = true;
        showError(validation.message);
        return null;
      }

      return {
        sheetName,
        visibleColumns: validation.sanitizedValue.split(","),
        viewName: `TVM_${AnsichtKey}`,
      };
    });

    if (hasErrors) {
      return;
    }

    // Filter out any null values from failed validations
    configs[AnsichtKey] = newConfigs.filter((config) => config !== null);

    await saveSettingsToStorage(configs);
    showSuccess("Einstellungen erfolgreich gespeichert!");
  } catch (error) {
    console.error("Fehler beim speichern der Einstellungen:", error);
    showError("Fehler beim speichern der Einstellungen");
  }
}

/**
 * Applies view settings for a specific Ansicht
 * @param {string} AnsichtKey - Key to identify Ansicht configuration
 */
async function applyAnsichtView(AnsichtKey) {
  try {
    const configs = await loadSettingsFromStorage();
    const AnsichtConfigs = configs[AnsichtKey];

    if (!AnsichtConfigs || AnsichtConfigs.length === 0) {
      showError("Keine Konfiguration für dieses Ansicht gefunden. Bitte konfigurieren Sie zuerst die Ansicht.");
      return;
    }

    await Excel.run(async (context) => {
      for (const config of AnsichtConfigs) {
        const sheet = context.workbook.worksheets.getItem(config.sheetName);

        // Load all named sheet views to check if the view alreadz exists
        const sheetViews = sheet.namedSheetViews;
        sheetViews.load("items/name");
        await context.sync();

        // Find existing view or create new one
        let currentView;
        const existingView = sheetViews.items.find((view) => view.name === config.viewName);

        if (existingView) {
          // If view exists, use it
          currentView = existingView;
        } else {
          // If view doesn't exist, create it
          currentView = sheetViews.add(config.viewName);
        }

        // Activate the view
        currentView.activate();
        await context.sync();

        // Get all columns in used range
        const usedRange = sheet.getUsedRange();
        usedRange.load("columnCount");
        await context.sync();

        // Hide all columns first
        for (let i = 1; i <= usedRange.columnCount; i++) {
          const columnLetter = getColumnLetter(i);
          try {
            sheet.getRange(`${columnLetter}:${columnLetter}`).columnHidden = true;
          } catch (error) {
            console.error(`Error hidding column (${columnLetter}:${columnLetter})`, error);
          }
        }

        // Show only the specified columns
        config.visibleColumns.forEach((colLetter) => {
          try {
            sheet.getRange(`${colLetter}:${colLetter}`).columnHidden = false;
          } catch (error) {
            console.error(`Error showing column (${colLetter}:${colLetter})`, error);
          }
        });

        await context.sync();
      }

      showSuccess("Ansicht erfolgreich angewendet!");
    });
  } catch (error) {
    if (error instanceof OfficeExtension.Error) {
      console.error("Office Extenstion Error:", JSON.stringify(error.stack));
      showError(
        "Fehler aufgetreten. Um Ansichten anwenden zu können, muss sich die Datei auf einem Sharepoint oder OneDrive befinden."
      );
    } else {
      console.error("Unknown Error:", error);
      showError("Unbekannter Fehler aufgetreten.");
    }
  }
}

/**
 * Helper function to convert column number to letter
 * @param {number} columnNumber - The column number to convert
 * @returns {string} The column letter
 */
function getColumnLetter(columnNumber) {
  let dividend = columnNumber;
  let columnName = "";
  let modulo;

  while (dividend > 0) {
    modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}

/**
 * Helper function to show messages to the user
 * @param {string} message - Message to display
 * @param {string} type - Message type ('success' or 'error')
 */
function showError(message, type = "error") {
  const messageDisplay = document.getElementById("messageDisplay");
  messageDisplay.textContent = message;
  messageDisplay.className = "message-display " + (type === "success" ? "message-success" : "message-error");
  messageDisplay.style.display = "block";

  setTimeout(() => {
    messageDisplay.style.display = "none";
  }, 3000);
}

// Delete selected Ansicht configuration
async function deleteAnsicht(event) {
  const AnsichtSelect = document.getElementById("AnsichtSelect");
  const selectedAnsicht = AnsichtSelect.value;

  if (selectedAnsicht) {
    const configs = await loadSettingsFromStorage();
    delete configs[selectedAnsicht];
    await saveSettingsToStorage(configs);
    updateAnsichtSelectOptions();
    toggleModal(event);
    showSuccess(`Ansicht "${selectedAnsicht}" wurde gelöscht.`);
  }
}

async function renameAnsicht() {
  document.getElementById("AnsichtManagement").style.display = "flex";
  document.getElementById("btnConfirmRenameAnsicht").style.display = "flex";
  document.getElementById("btnConfirmAddAnsicht").style.display = "none";
}

// Function to handle initiating Ansicht rename
async function confirmRenameAnsicht() {
  const AnsichtSelect = document.getElementById("AnsichtSelect");
  const AnsichtInput = document.getElementById("AnsichtInput").value.trim();

  const oldAnsichtName = AnsichtSelect.value;

  const configs = await loadSettingsFromStorage();

  if (!configs.hasOwnProperty(AnsichtInput)) {
    configs[AnsichtInput] = configs[oldAnsichtName];
    delete configs[oldAnsichtName];
    await saveSettingsToStorage(configs);
    document.getElementById("AnsichtManagement").style.display = "none";
    updateAnsichtSelectOptions();
    showSuccess(`Ansicht wurde in "${AnsichtInput}" umbenannt.`);
  } else {
    showError("Ein Ansicht mit diesem Namen existiert bereits.");
    return;
  }
}

async function addAnsicht() {
  document.getElementById("AnsichtManagement").style.display = "flex";
  document.getElementById("btnConfirmRenameAnsicht").style.display = "none";
  document.getElementById("btnConfirmAddAnsicht").style.display = "flex";
  document.getElementById("AnsichtInput").value = "";
}

// Confirm adding or renaming Ansicht
async function confirmAddAnsicht() {
  const AnsichtInput = document.getElementById("AnsichtInput").value.trim();
  const configs = await loadSettingsFromStorage();

  if (!AnsichtInput) {
    showError("Bitte geben Sie einen gültigen Ansichtnamen ein.");
    return;
  }

  if (!configs.hasOwnProperty(AnsichtInput)) {
    configs[AnsichtInput] = [];
    await saveSettingsToStorage(configs);
    document.getElementById("AnsichtManagement").style.display = "none";
    updateAnsichtSelectOptions();
    showSuccess(`Ansicht "${AnsichtInput}" wurde hinzugefügt.`);
  } else {
    showError("Dieses Ansicht existiert bereits.");
    return;
  }
}

// Update the Ansicht select dropdown options
async function updateAnsichtSelectOptions() {
  const AnsichtSelect = document.getElementById("AnsichtSelect");
  const AnsichtButtonsContainer = document.getElementById("AnsichtButtonsContainer");
  const configs = await loadSettingsFromStorage();

  // Clear existing options and buttons
  AnsichtSelect.innerHTML = "";
  AnsichtButtonsContainer.innerHTML = "";

  Object.keys(configs).forEach((Ansicht) => {
    // Populate dropdown
    const option = document.createElement("option");
    option.value = Ansicht;
    option.textContent = Ansicht;
    AnsichtSelect.appendChild(option);

    // Create Ansicht button
    const button = document.createElement("button");
    button.id = `btn${Ansicht}`;
    button.className = "button is-view";
    button.textContent = Ansicht;
    button.onclick = () => applyAnsichtView(Ansicht);
    AnsichtButtonsContainer.appendChild(button);
  });
}

/**
 * Helper function to show success messages
 * @param {string} message - Success message to display
 */
function showSuccess(message) {
  showError(message, "success");
}

// Toggle modal
const toggleModal = (event) => {
  event.preventDefault();
  const modal = document.getElementById(event.currentTarget.dataset.target);
  if (!modal) return;
  modal && (modal.open ? closeModal(modal) : openModal(modal));
};

// Open modal
const openModal = (modal) => {
  const { documentElement: html } = document;
  const scrollbarWidth = getScrollbarWidth();
  if (scrollbarWidth) {
    html.style.setProperty(scrollbarWidthCssVar, `${scrollbarWidth}px`);
  }
  html.classList.add(isOpenClass, openingClass);
  setTimeout(() => {
    visibleModal = modal;
    html.classList.remove(openingClass);
  }, animationDuration);
  modal.showModal();
};

// Close modal
const closeModal = (modal) => {
  visibleModal = null;
  const { documentElement: html } = document;
  html.classList.add(closingClass);
  setTimeout(() => {
    html.classList.remove(closingClass, isOpenClass);
    html.style.removeProperty(scrollbarWidthCssVar);
    modal.close();
  }, animationDuration);
};

// Close with a click outside
document.addEventListener("click", (event) => {
  if (visibleModal === null) return;
  const modalContent = visibleModal.querySelector("article");
  const isClickInside = modalContent.contains(event.target);
  !isClickInside && closeModal(visibleModal);
});

// Close with Esc key
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && visibleModal) {
    closeModal(visibleModal);
  }
});

// Get scrollbar width
const getScrollbarWidth = () => {
  const scrollbarWidth = window.innerWidth - document.documentElement.clientWidth;
  return scrollbarWidth;
};

// Is scrollbar visible
const isScrollbarVisible = () => {
  return document.body.scrollHeight > screen.height;
};
