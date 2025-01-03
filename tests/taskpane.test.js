// Mock the Office object
global.Office = {
  onReady: jest.fn((callback) => {
    callback({ host: "Excel" }); // Simulate that the host is Excel
  }),
  HostType: {
    Excel: "Excel", // Define the Excel host type
  },
  context: {
    document: {
      settings: {
        get: jest.fn(),
        set: jest.fn(),
        saveAsync: jest.fn((callback) => {
          callback({ status: "succeeded" }); // Simulate successful save
        }),
      },
    },
  },
};

// Functions to be tested
const {
  convertColumnLetterToNumber,
  validateSingleColumn,
  sanitizeColumnInput,
  validateColumnInput,
} = require("../src/taskpane/taskpane");

describe("Excel Column Utilities", () => {
  describe("convertColumnLetterToNumber", () => {
    test("should convert single letter A to 1", () => {
      expect(convertColumnLetterToNumber("A")).toBe(1);
    });

    test("should convert multiple letters AA to 27", () => {
      expect(convertColumnLetterToNumber("AA")).toBe(27);
    });

    test("should convert last column XFD to 16384", () => {
      expect(convertColumnLetterToNumber("XFD")).toBe(16384);
    });

    test("should handle invalid input gracefully", () => {
      expect(() => convertColumnLetterToNumber("")).toThrow();
      expect(() => convertColumnLetterToNumber("1")).toThrow();
      expect(() => convertColumnLetterToNumber("A1")).toThrow();
    });
  });

  describe("validateSingleColumn", () => {
    test("should validate valid column A", () => {
      const result = validateSingleColumn("A");
      expect(result.isValid).toBe(true);
    });

    test("should validate valid column AA", () => {
      const result = validateSingleColumn("AA");
      expect(result.isValid).toBe(true);
    });

    test("should invalidate column longer than 3 characters", () => {
      const result = validateSingleColumn("AAAA");
      expect(result.isValid).toBe(false);
      expect(result.message).toContain("maximal 3 Buchstaben lang");
    });

    test("should invalidate column beyond Excel limits", () => {
      const result = validateSingleColumn("XFE");
      expect(result.isValid).toBe(false);
      expect(result.message).toContain("liegt außerhalb des gültigen Excel-Bereichs");
    });
  });

  describe("sanitizeColumnInput", () => {
    test("should remove whitespace and convert to uppercase", () => {
      expect(sanitizeColumnInput(" a , b , c ")).toBe("A,B,C");
    });

    test("should remove invalid characters", () => {
      expect(sanitizeColumnInput("A,B,C,1")).toBe("A,B,C");
    });

    test("should handle multiple commas", () => {
      expect(sanitizeColumnInput("A,,B,,C")).toBe("A,B,C");
    });

    test("should return empty string for invalid input", () => {
      expect(sanitizeColumnInput("123")).toBe("");
    });
  });

  describe("validateColumnInput", () => {
    test("should validate correct input format", () => {
      const result = validateColumnInput("A,B,C");
      expect(result.isValid).toBe(true);
      expect(result.sanitizedValue).toBe("A,B,C");
      expect(result.columns).toEqual(["A", "B", "C"]);
    });

    test("should invalidate empty input", () => {
      const result = validateColumnInput("");
      expect(result.isValid).toBe(false);
      expect(result.message).toBe("Bitte geben Sie mindestens eine Spalte an.");
    });

    test("should invalidate input with invalid characters", () => {
      const result = validateColumnInput("A,B,C,1");
      expect(result.isValid).toBe(false);
      expect(result.message).toContain("Ungültiges Format");
    });

    test("should invalidate input with empty column references", () => {
      const result = validateColumnInput("A,,C");
      expect(result.isValid).toBe(false);
      expect(result.message).toBe("Leere Spaltenreferenz gefunden. Bitte überprüfen Sie die Kommas");
    });

    test("should validate input with valid columns", () => {
      const result = validateColumnInput("A,B,C");
      expect(result.isValid).toBe(true);
    });

    test("should handle input with columns exceeding limits", () => {
      const result = validateColumnInput("A,B,AAAA");
      expect(result.isValid).toBe(false);
      expect(result.message).toContain("Ungültige Spaltenreferenz");
    });
  });
});
