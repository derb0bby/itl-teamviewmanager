// The property key where we'll store our team configurations
const TEAM_CONFIGS_PROPERTY = "TeamViewConfigurations";

/**
 * Validates column input format (e.g., "A,B,C")
 * @param {string} input - The column input string to validate
 * @returns {Object} Object containing validation result and error message
 */
function validateColumnInput(input) {
  if (!input || input.trim() === "") {
    return {
      isValid: false,
      message: "Bitte geben Sie mindestens eine Spalte ein.",
    };
  }

  // Remove any whitespace and convert to uppercase
  const sanitizedInput = sanitizeColumnInput(input);

  // Check if the input contains only letters and commas
  const validFormat = /^[A-Z]+(,[A-Z]+)*$/.test(sanitizedInput);

  if (!validFormat) {
    return {
      isValid: false,
      message: "Ungültiges Format. Bitte verwenden Sie nur Buchstaben und Kommas (A,B,C).",
    };
  }

  return {
    isValid: true,
    sanitizedValue: sanitizedInput,
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
    .replace(/[^A-Z,]/g, ""); // Remove any characters that aren't letters or commas
}

// Initialize Office JS
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Set up event handlers for existing buttons
    document.getElementById("btnTeamLager").onclick = () => applyTeamView("Team_Lager");
    document.getElementById("btnTeamSM").onclick = () => applyTeamView("Team_SM");
    document.getElementById("btnTeamTechnik").onclick = () => applyTeamView("Team_Technik");
    document.getElementById("btnConfigureView").onclick = showConfigurationDialog;
    document.getElementById("btnCloseDialog").onclick = hideConfigurationDialog;
    document.getElementById("btnSaveConfig").onclick = saveConfiguration;
    document.getElementById("btnAddSheet").onclick = addSheetConfiguration;
    document.getElementById("teamSelect").onchange = loadTeamConfiguration;
  }
});

/**
 * Shows the configuration dialog and loads current settings
 */
async function showConfigurationDialog() {
  const dialog = document.getElementById("configDialog");
  dialog.style.display = "flex";

  // Load current team's configuration
  const teamSelect = document.getElementById("teamSelect");
  await loadTeamConfiguration(teamSelect.value);
}

/**
 * Hides the configuration dialog
 */
function hideConfigurationDialog() {
  document.getElementById("configDialog").style.display = "none";
}

/**
 * Loads configurations from workbook custom properties
 * @returns {Promise<Object>} The team configurations
 */
/**
 * Loads configurations from the Office.context.document.settings property bag.
 * Returns the stored configuration object or default values if none exist.
 */
async function loadSettingsFromStorage() {
  try {
    const storedConfig = Office.context.document.settings.get(TEAM_CONFIGS_PROPERTY);

    if (!storedConfig) {
      // Return default empty configuration if no configuration exists
      return {
        Team_Lager: [],
        Team_SM: [],
        Team_Technik: [],
      };
    }

    return JSON.parse(storedConfig); // Parse and return the stored configuration
  } catch (error) {
    console.error("Error loading configurations:", error);
    showError("Error loading configurations");
    return {
      Team_Lager: [],
      Team_SM: [],
      Team_Technik: [],
    };
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
    Office.context.document.settings.set(TEAM_CONFIGS_PROPERTY, jsonString);

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
 * Loads configuration for a specific team
 */
async function loadTeamConfiguration() {
  const teamKey = document.getElementById("teamSelect").value;
  const configs = await loadSettingsFromStorage();
  const teamConfig = configs[teamKey] || [];

  // Get available sheets
  const sheets = await getWorksheetNames();

  // Clear existing configuration UI
  const sheetConfig = document.getElementById("sheetConfig");
  sheetConfig.innerHTML = "";

  // Create configuration UI for each sheet
  teamConfig.forEach((config, index) => {
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
    <select class="sheet-select full-width">
      ${sheets
        .map(
          (sheet) =>
            `<option value="${sheet}" ${config && config.sheetName === sheet ? "selected" : ""}>${sheet}</option>`
        )
        .join("")}
    </select>
    <input type="text" class="columns-input" placeholder="Angezeigte Spalten (z.B., A,C,E)" 
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
    const teamKey = document.getElementById("teamSelect").value;
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
        viewName: `TVM_${teamKey}`,
      };
    });

    if (hasErrors) {
      return;
    }

    // Filter out any null values from failed validations
    configs[teamKey] = newConfigs.filter((config) => config !== null);

    await saveSettingsToStorage(configs);
    showSuccess("Einstellungen erfolgreich gespeichert!");
  } catch (error) {
    console.error("Fehler beim speichern der Einstellungen:", error);
    showError("Fehler beim speichern der Einstellungen");
  }
}

/**
 * Applies view settings for a specific team
 * @param {string} teamKey - Key to identify team configuration
 */
async function applyTeamView(teamKey) {
  try {
    const configs = await loadSettingsFromStorage();
    const teamConfigs = configs[teamKey];

    if (!teamConfigs || teamConfigs.length === 0) {
      showError("Keine Konfiguration für dieses Team gefunden. Bitte konfigurieren Sie zuerst die Ansicht.");
      return;
    }

    await Excel.run(async (context) => {
      for (const config of teamConfigs) {
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

/**
 * Helper function to show success messages
 * @param {string} message - Success message to display
 */
function showSuccess(message) {
  showError(message, "success");
}
