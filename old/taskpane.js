// The property key where we'll store our team configurations
const TEAM_CONFIGS_PROPERTY = "TeamViewConfigurations";

// Initialize Office JS
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Set up event handlers for existing buttons
    document.getElementById("btnTeamLager").onclick = () => applyTeamView("teamLager");
    document.getElementById("btnTeamSM").onclick = () => applyTeamView("teamSM");
    document.getElementById("btnTeamTechnik").onclick = () => applyTeamView("teamTechnik");
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
async function loadConfigurations() {
  try {
    return await Excel.run(async (context) => {
      const customProperties = context.workbook.properties.custom;
      const configProperty = customProperties.getItemOrNullObject(TEAM_CONFIGS_PROPERTY);
      configProperty.load("value");
      await context.sync();

      if (configProperty.isNullObject) {
        // Return default empty configuration
        return {
          teamLager: [],
          teamSM: [],
          teamTechnik: [],
        };
      }

      return JSON.parse(configProperty.value);
    });
  } catch (error) {
    console.error("Error loading configurations:", error);
    showError("Error loading configurations");
    return {
      teamLager: [],
      teamSM: [],
      teamTechnik: [],
    };
  }
}

/**
 * Saves configurations to workbook custom properties
 * @param {Object} configs - The configurations to save
 */
async function saveConfigurations(configs) {
  try {
    await Excel.run(async (context) => {
      const customProperties = context.workbook.properties.custom;

      // Try to get existing property
      const existingProperty = customProperties.getItemOrNullObject(TEAM_CONFIGS_PROPERTY);
      await context.sync();

      if (!existingProperty.isNullObject) {
        // If property exists, delete it first
        existingProperty.delete();
      }

      // Add new property with updated configurations
      customProperties.add(TEAM_CONFIGS_PROPERTY, JSON.stringify(configs));
      await context.sync();

      // Close Windows
      hideConfigurationDialog();
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
  const configs = await loadConfigurations();
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
 * Creates UI elements for sheet configuration
 * @param {HTMLElement} container - Container element
 * @param {string[]} sheets - Available worksheet names
 * @param {Object} config - Existing configuration
 * @param {number} index - Configuration index
 */
function addSheetConfigurationUI(container, sheets, config, index) {
  const div = document.createElement("div");
  div.className = "sheet-config";
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
               value="${config ? config.visibleColumns.join(",") : ""}">
        <button class="button-remove" onclick="removeSheetConfig(${index})">Löschen</button>
    `;
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
    const configs = await loadConfigurations();

    configs[teamKey] = Array.from(sheetConfigs).map((config) => {
      const sheetName = config.querySelector(".sheet-select").value;
      const visibleColumns = config
        .querySelector(".columns-input")
        .value.split(",")
        .map((col) => col.trim())
        .filter((col) => col);

      return {
        sheetName,
        visibleColumns,
        viewName: `Team_${teamKey}`,
      };
    });

    await saveConfigurations(configs);
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
    const configs = await loadConfigurations();
    const teamConfigs = configs[teamKey];

    if (!teamConfigs || teamConfigs.length === 0) {
      showError("No configuration found for this team. Please configure the view first.");
      return;
    }

    await Excel.run(async (context) => {
      for (const config of teamConfigs) {
        const sheet = context.workbook.worksheets.getItem(config.sheetName);

        // Load all named sheet views to check if our view exists
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
        console.log(usedRange.columnCount);
        for (let i = 1; i <= usedRange.columnCount; i++) {
          const columnLetter = getColumnLetter(i);
          sheet.getRange(`${columnLetter}:${columnLetter}`).columnHidden = true;
        }

        // Show only the specified columns
        config.visibleColumns.forEach((colLetter) => {
          sheet.getRange(`${colLetter}:${colLetter}`).columnHidden = false;
        });

        await context.sync();
      }

      showSuccess("View applied successfully!");
    });
  } catch (error) {
    console.error("Error:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", JSON.stringify(error.debugInfo));
    }
    showError("An error occurred while applying the view");
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
