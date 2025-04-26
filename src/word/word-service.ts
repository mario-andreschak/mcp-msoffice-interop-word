import winax from 'winax';
const { Object: WinaxObject } = winax; // Destructure and rename Object

// Basic interface for Word Application object (replace with more specific types later if possible)
interface WordApplication {
  Documents: any; // Word.Documents collection
  ActiveDocument: any; // Word.Document
  Visible: boolean;
  Quit(SaveChanges?: any, OriginalFormat?: any, RouteDocument?: any): void;
  // Add other necessary properties and methods
}

// Basic interface for Word Document object
interface WordDocument {
  Save(): void;
  SaveAs2(FileName?: any, FileFormat?: any, LockComments?: any, Password?: any, AddToRecentFiles?: any, WritePassword?: any, ReadOnlyRecommended?: any, EmbedTrueTypeFonts?: any, SaveNativePictureFormat?: any, SaveFormsData?: any, SaveAsAOCELetter?: any, Encoding?: any, InsertLineBreaks?: any, AllowSubstitutions?: any, LineEnding?: any, AddBiDiMarks?: any, CompatibilityMode?: any): void;
  Close(SaveChanges?: any, OriginalFormat?: any, RouteDocument?: any): void;
  Content: any; // Word.Range
  Paragraphs: any; // Word.Paragraphs
  Tables: any; // Word.Tables
  InlineShapes: any; // Word.InlineShapes
  Shapes: any; // Word.Shapes
  Sections: any; // Word.Sections
  ActiveWindow: any; // Word.Window
  PageSetup: any; // Word.PageSetup
  // Add other necessary properties and methods
}

class WordService {
  private wordApp: WordApplication | null = null;

  /**
   * Gets the currently running Word application instance or creates a new one.
   * Ensures Word is visible.
   */
  public async getWordApplication(): Promise<WordApplication> {
    if (this.wordApp) {
      try {
        // More comprehensive check for instance validity
        // Check basic property access
        this.wordApp.Visible;
        
        // Check if Documents collection is accessible and valid
        const docs = this.wordApp.Documents;
        // Try to access a property of Documents to ensure it's fully valid
        const count = docs.Count;
        
        return this.wordApp;
      } catch (error) {
        console.warn("Existing Word instance seems invalid, creating a new one.", error);
        this.wordApp = null; // Reset if invalid
      }
    }

    try {
      console.log("Attempting to get or create Word.Application instance...");
      // Try to get an existing instance first, then create if not found
      // Use the destructured WinaxObject constructor
      this.wordApp = new WinaxObject("Word.Application", { activate: true }) as WordApplication;
      this.wordApp.Visible = true; // Make sure Word is visible for interaction
      console.log("Word.Application instance obtained successfully.");
      return this.wordApp;
    } catch (error) {
      console.error("Failed to get or create Word.Application instance:", error);
      throw new Error(`Failed to initialize Word.Application. Make sure Microsoft Word is installed. Error: ${error}`);
    }
  }

  /**
   * Gets the active document in the Word application.
   * Throws an error if Word is not running or no document is active.
   */
  public async getActiveDocument(): Promise<WordDocument> {
    const app = await this.getWordApplication();
    try {
      const activeDoc = app.ActiveDocument;
      if (!activeDoc) {
        throw new Error("No active document found in Word.");
      }
      return activeDoc as WordDocument;
    } catch (error) {
      console.error("Failed to get active document:", error);
      throw new Error(`Failed to get active document. Error: ${error}`);
    }
  }

  // --- Document Methods ---

  /**
   * Creates a new Word document.
   */
  public async createDocument(): Promise<WordDocument> {
    const app = await this.getWordApplication();
    try {
      const newDoc = app.Documents.Add();
      return newDoc as WordDocument;
    } catch (error) {
      console.error("Failed to create new document:", error);
      throw new Error(`Failed to create new document. Error: ${error}`);
    }
  }

  /**
   * Opens an existing Word document.
   * @param filePath The path to the document file.
   */
  public async openDocument(filePath: string): Promise<WordDocument> {
    const app = await this.getWordApplication();
    try {
      // Ensure the path is absolute and correctly formatted if needed
      const openedDoc = app.Documents.Open(filePath);
      return openedDoc as WordDocument;
    } catch (error) {
      console.error(`Failed to open document at path: ${filePath}`, error);
      throw new Error(`Failed to open document: ${filePath}. Error: ${error}`);
    }
  }

   /**
   * Saves the active document.
   */
    public async saveActiveDocument(): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
        doc.Save();
      } catch (error) {
        console.error("Failed to save active document:", error);
        throw new Error(`Failed to save active document. Error: ${error}`);
      }
    }

  /**
   * Saves the active document with a new name or format.
   * @param filePath The new path for the document.
   * @param fileFormat Optional Word save format constant (e.g., WdSaveFormat.wdFormatDocumentDefault).
   */
  public async saveActiveDocumentAs(filePath: string, fileFormat?: any): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      // WdSaveFormat enumeration needs to be accessible or defined
      // Example: wdFormatDocumentDefault = 16
      const format = fileFormat ?? 16; // Default to .docx
      doc.SaveAs2(filePath, format);
    } catch (error) {
      console.error(`Failed to save document as: ${filePath}`, error);
      throw new Error(`Failed to save document as: ${filePath}. Error: ${error}`);
    }
  }

  /**
   * Closes the specified document.
   * @param doc The document object to close.
   * @param saveChanges Optional WdSaveOptions constant (e.g., WdSaveOptions.wdDoNotSaveChanges).
   */
  public async closeDocument(doc: WordDocument, saveChanges?: any): Promise<void> {
    try {
      // WdSaveOptions enumeration needs to be accessible or defined
      // Example: wdDoNotSaveChanges = 0, wdPromptToSaveChanges = -2, wdSaveChanges = -1
      const saveOpt = saveChanges ?? 0; // Default to not saving changes
      doc.Close(saveOpt);
    } catch (error) {
      console.error("Failed to close document:", error);
      // Avoid throwing if close fails, might already be closed or Word unresponsive
      console.warn(`Could not definitively close document. Error: ${error}`);
    }
  }


  /**
   * Quits the Word application.
   * Handles potential errors during quit.
   */
  public async quitWord(): Promise<void> {
    if (this.wordApp) {
      try {
        // WdSaveOptions enumeration
        // Example: wdDoNotSaveChanges = 0
        this.wordApp.Quit(0); // Quit without saving changes
        this.wordApp = null; // Clear the reference
        console.log("Word application quit successfully.");
      } catch (error) {
        console.error("Error quitting Word application:", error);
        // Don't re-throw, as Word might already be closed or unresponsive
        this.wordApp = null; // Clear reference even on error
      }
    } else {
        console.log("Word application instance not found, nothing to quit.");
    }
  }

  // --- Text Manipulation Methods ---

  /**
   * Inserts text at the current selection point.
   * @param text The text to insert.
   */
  public async insertText(text: string): Promise<void> {
    const app = await this.getWordApplication();
    try {
      app.ActiveDocument.ActiveWindow.Selection.TypeText(text);
    } catch (error) {
      console.error("Failed to insert text:", error);
      throw new Error(`Failed to insert text. Error: ${error}`);
    }
  }

  /**
   * Deletes the current selection or a specified number of characters.
   * @param count Number of characters to delete (default: 1). Positive deletes forward, negative deletes backward.
   * @param unit The unit to delete (default: character). Use WdUnits enum values (e.g., 1 for character, 2 for word).
   */
  public async deleteText(count: number = 1, unit: number = 1 /* wdCharacter */): Promise<void> {
      const app = await this.getWordApplication();
      try {
          // WdUnits enumeration: wdCharacter = 1, wdWord = 2, etc.
          // Positive count deletes forward, negative count deletes backward from the start of the selection.
          // If selection is collapsed, positive deletes after insertion point, negative deletes before.
          if (count > 0) {
              app.ActiveDocument.ActiveWindow.Selection.Delete(unit, count);
          } else if (count < 0) {
              // Move start back and then delete forward
              app.ActiveDocument.ActiveWindow.Selection.MoveStart(unit, count); // Move start back
              app.ActiveDocument.ActiveWindow.Selection.Delete(unit, Math.abs(count)); // Delete forward
          }
          // If count is 0, do nothing
      } catch (error) {
          console.error("Failed to delete text:", error);
          throw new Error(`Failed to delete text. Error: ${error}`);
      }
  }


  /**
   * Finds and replaces text in the active document.
   * @param findText Text to find.
   * @param replaceText Text to replace with.
   * @param matchCase Match case sensitivity.
   * @param matchWholeWord Match whole words only.
   * @param replaceAll Replace all occurrences or just the first.
   */
   public async findAndReplace(
    findText: string,
    replaceText: string,
    matchCase: boolean = false,
    matchWholeWord: boolean = false,
    replaceAll: boolean = true
  ): Promise<boolean> {
    const doc = await this.getActiveDocument();
    try {
      const find = doc.Content.Find;
      find.ClearFormatting(); // Clear previous find formatting
      find.Replacement.ClearFormatting(); // Clear previous replacement formatting

      find.Text = findText;
      find.Replacement.Text = replaceText;
      find.Forward = true;
      find.Wrap = 1; // wdFindContinue
      find.Format = false;
      find.MatchCase = matchCase;
      find.MatchWholeWord = matchWholeWord;
      find.MatchWildcards = false;
      find.MatchSoundsLike = false;
      find.MatchAllWordForms = false;

      // WdReplace enumeration: wdReplaceNone = 0, wdReplaceOne = 1, wdReplaceAll = 2
      const replaceOption = replaceAll ? 2 : 1;

      const found = find.Execute(undefined, undefined, undefined, undefined, undefined,
                                 undefined, undefined, undefined, undefined, undefined,
                                 replaceOption); // Execute the find and replace

      return found; // Returns true if text was found and replaced (or just found if replaceOption is wdReplaceNone)
    } catch (error) {
      console.error("Failed to find and replace text:", error);
      throw new Error(`Failed to find and replace text. Error: ${error}`);
    }
  }

  /**
   * Toggles bold formatting for the current selection.
   */
  public async toggleBold(): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const font = app.ActiveDocument.ActiveWindow.Selection.Font;
      // wdToggle = 9999998
      font.Bold = 9999998;
    } catch (error) {
      console.error("Failed to toggle bold:", error);
      throw new Error(`Failed to toggle bold formatting. Error: ${error}`);
    }
  }

  /**
   * Toggles italic formatting for the current selection.
   */
  public async toggleItalic(): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const font = app.ActiveDocument.ActiveWindow.Selection.Font;
      // wdToggle = 9999998
      font.Italic = 9999998;
    } catch (error) {
      console.error("Failed to toggle italic:", error);
      throw new Error(`Failed to toggle italic formatting. Error: ${error}`);
    }
  }

  /**
   * Toggles underline formatting for the current selection.
   * @param underlineStyle Optional WdUnderline value (e.g., 1 for single underline). Default toggles single underline.
   */
  public async toggleUnderline(underlineStyle: number = 1 /* wdUnderlineSingle */): Promise<void> {
      const app = await this.getWordApplication();
      try {
          const font = app.ActiveDocument.ActiveWindow.Selection.Font;
          // WdUnderline enumeration: wdUnderlineNone = 0, wdUnderlineSingle = 1, etc.
          // wdToggle = 9999998
          if (font.Underline === underlineStyle) {
              font.Underline = 0; // wdUnderlineNone - Turn off if it's already the specified style
          } else {
              font.Underline = underlineStyle; // Apply the specified style
          }
      } catch (error) {
          console.error("Failed to toggle underline:", error);
          throw new Error(`Failed to toggle underline formatting. Error: ${error}`);
      }
  }

  // --- Paragraph Formatting Methods ---

  /**
   * Sets the alignment for the selected paragraphs.
   * @param alignment Alignment type (WdParagraphAlignment enum value: 0=Left, 1=Center, 2=Right, 3=Justify).
   */
  public async setParagraphAlignment(alignment: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      // WdParagraphAlignment: wdAlignParagraphLeft = 0, wdAlignParagraphCenter = 1, wdAlignParagraphRight = 2, wdAlignParagraphJustify = 3
      app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.Alignment = alignment;
    } catch (error) {
      console.error("Failed to set paragraph alignment:", error);
      throw new Error(`Failed to set paragraph alignment. Error: ${error}`);
    }
  }

  /**
   * Sets the left indent for the selected paragraphs.
   * @param indentPoints Indentation value in points.
   */
  public async setParagraphLeftIndent(indentPoints: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.LeftIndent = indentPoints;
    } catch (error) {
      console.error("Failed to set left indent:", error);
      throw new Error(`Failed to set left indent. Error: ${error}`);
    }
  }

  /**
   * Sets the right indent for the selected paragraphs.
   * @param indentPoints Indentation value in points.
   */
  public async setParagraphRightIndent(indentPoints: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.RightIndent = indentPoints;
    } catch (error) {
      console.error("Failed to set right indent:", error);
      throw new Error(`Failed to set right indent. Error: ${error}`);
    }
  }

    /**
   * Sets the first line indent for the selected paragraphs.
   * @param indentPoints Indentation value in points (positive for indent, negative for hanging indent).
   */
    public async setParagraphFirstLineIndent(indentPoints: number): Promise<void> {
        const app = await this.getWordApplication();
        try {
            app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.FirstLineIndent = indentPoints;
        } catch (error) {
            console.error("Failed to set first line indent:", error);
            throw new Error(`Failed to set first line indent. Error: ${error}`);
        }
    }

  /**
   * Sets the space before the selected paragraphs.
   * @param spacePoints Space value in points.
   */
  public async setParagraphSpaceBefore(spacePoints: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.SpaceBefore = spacePoints;
    } catch (error) {
      console.error("Failed to set space before:", error);
      throw new Error(`Failed to set space before paragraph. Error: ${error}`);
    }
  }

  /**
   * Sets the space after the selected paragraphs.
   * @param spacePoints Space value in points.
   */
  public async setParagraphSpaceAfter(spacePoints: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.SpaceAfter = spacePoints;
    } catch (error) {
      console.error("Failed to set space after:", error);
      throw new Error(`Failed to set space after paragraph. Error: ${error}`);
    }
  }

  /**
   * Sets the line spacing for the selected paragraphs.
   * @param lineSpacingRule WdLineSpacing enum value (0=Single, 1=1.5 lines, 2=Double, 3=AtLeast, 4=Exactly, 5=Multiple).
   * @param lineSpacingValue Value for AtLeast, Exactly, or Multiple spacing (in points for AtLeast/Exactly, multiplier for Multiple).
   */
  public async setParagraphLineSpacing(lineSpacingRule: number, lineSpacingValue?: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      // WdLineSpacing: wdLineSpaceSingle = 0, wdLineSpace1pt5 = 1, wdLineSpaceDouble = 2,
      // wdLineSpaceAtLeast = 3, wdLineSpaceExactly = 4, wdLineSpaceMultiple = 5
      const paraFormat = app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat;
      paraFormat.LineSpacingRule = lineSpacingRule;
      if (lineSpacingValue !== undefined && lineSpacingRule >= 3) { // Only set LineSpacing if rule requires it
        paraFormat.LineSpacing = lineSpacingValue;
      }
    } catch (error) {
      console.error("Failed to set line spacing:", error);
      throw new Error(`Failed to set line spacing. Error: ${error}`);
    }
  }

  // --- Table Methods ---

  /**
   * Adds a table at the current selection.
   * @param numRows Number of rows.
   * @param numCols Number of columns.
   * @param defaultTableBehavior Optional WdDefaultTableBehavior value.
   * @param autoFitBehavior Optional WdAutoFitBehavior value.
   */
  public async addTable(numRows: number, numCols: number, defaultTableBehavior?: number, autoFitBehavior?: number): Promise<any /* Word.Table */> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      // WdDefaultTableBehavior: wdWord8TableBehavior = 0, wdWord9TableBehavior = 1
      // WdAutoFitBehavior: wdAutoFitFixed = 0, wdAutoFitContent = 1, wdAutoFitWindow = 2
      const table = selection.Tables.Add(selection.Range, numRows, numCols, defaultTableBehavior, autoFitBehavior);
      return table;
    } catch (error) {
      console.error("Failed to add table:", error);
      throw new Error(`Failed to add table. Error: ${error}`);
    }
  }

  /**
   * Gets a specific cell in a table.
   * @param tableIndex Index of the table in the document (1-based).
   * @param rowIndex Row index (1-based).
   * @param colIndex Column index (1-based).
   */
  public async getTableCell(tableIndex: number, rowIndex: number, colIndex: number): Promise<any /* Word.Cell */> {
    const doc = await this.getActiveDocument();
    try {
      if (tableIndex <= 0 || tableIndex > doc.Tables.Count) {
        throw new Error(`Table index ${tableIndex} is out of bounds.`);
      }
      const table = doc.Tables.Item(tableIndex);
      const cell = table.Cell(rowIndex, colIndex);
      return cell;
    } catch (error) {
      console.error(`Failed to get cell (${rowIndex}, ${colIndex}) in table ${tableIndex}:`, error);
      throw new Error(`Failed to get cell. Error: ${error}`);
    }
  }

  /**
   * Sets the text in a specific table cell.
   * @param tableIndex Index of the table in the document (1-based).
   * @param rowIndex Row index (1-based).
   * @param colIndex Column index (1-based).
   * @param text Text to set.
   */
  public async setTableCellText(tableIndex: number, rowIndex: number, colIndex: number, text: string): Promise<void> {
    try {
      const cell = await this.getTableCell(tableIndex, rowIndex, colIndex);
      cell.Range.Text = text;
    } catch (error) {
      console.error(`Failed to set text in cell (${rowIndex}, ${colIndex}) of table ${tableIndex}:`, error);
      throw new Error(`Failed to set cell text. Error: ${error}`);
    }
  }

  /**
   * Inserts a row in a table.
   * @param tableIndex Index of the table (1-based).
   * @param beforeRowIndex Optional index of the row to insert before (1-based). If omitted, adds to the end.
   */
  public async insertTableRow(tableIndex: number, beforeRowIndex?: number): Promise<any /* Word.Row */> {
      const doc = await this.getActiveDocument();
      try {
          if (tableIndex <= 0 || tableIndex > doc.Tables.Count) {
              throw new Error(`Table index ${tableIndex} is out of bounds.`);
          }
          const table = doc.Tables.Item(tableIndex);
          let newRow;
          if (beforeRowIndex !== undefined) {
              if (beforeRowIndex <= 0 || beforeRowIndex > table.Rows.Count + 1) { // Allow inserting after last row
                 throw new Error(`Row index ${beforeRowIndex} is out of bounds for insertion.`);
              }
              const refRow = beforeRowIndex <= table.Rows.Count ? table.Rows.Item(beforeRowIndex) : undefined;
              newRow = table.Rows.Add(refRow); // Inserts before refRow if provided, otherwise adds at end
          } else {
              newRow = table.Rows.Add(); // Add to the end
          }
          return newRow;
      } catch (error) {
          console.error(`Failed to insert row into table ${tableIndex}:`, error);
          throw new Error(`Failed to insert table row. Error: ${error}`);
      }
  }

  /**
   * Inserts a column in a table.
   * @param tableIndex Index of the table (1-based).
   * @param beforeColIndex Optional index of the column to insert before (1-based). If omitted, adds to the right end.
   */
  public async insertTableColumn(tableIndex: number, beforeColIndex?: number): Promise<any /* Word.Column */> {
      const doc = await this.getActiveDocument();
      try {
          if (tableIndex <= 0 || tableIndex > doc.Tables.Count) {
              throw new Error(`Table index ${tableIndex} is out of bounds.`);
          }
          const table = doc.Tables.Item(tableIndex);
           let newCol;
          if (beforeColIndex !== undefined) {
               if (beforeColIndex <= 0 || beforeColIndex > table.Columns.Count + 1) { // Allow inserting after last col
                 throw new Error(`Column index ${beforeColIndex} is out of bounds for insertion.`);
              }
              const refCol = beforeColIndex <= table.Columns.Count ? table.Columns.Item(beforeColIndex) : undefined;
              newCol = table.Columns.Add(refCol); // Inserts before refCol if provided, otherwise adds at end
          } else {
              newCol = table.Columns.Add(); // Add to the end (right)
          }
          return newCol;
      } catch (error) {
          console.error(`Failed to insert column into table ${tableIndex}:`, error);
          throw new Error(`Failed to insert table column. Error: ${error}`);
      }
  }

  /**
   * Applies an auto format style to a table.
   * @param tableIndex Index of the table (1-based).
   * @param formatName Name of the table style or a WdTableFormat enum value.
   * @param applyFormatting Optional flags for which parts of the format to apply (WdTableFormatApply enum values).
   */
  public async applyTableAutoFormat(tableIndex: number, formatName: string | number, applyFormatting?: number): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
          if (tableIndex <= 0 || tableIndex > doc.Tables.Count) {
              throw new Error(`Table index ${tableIndex} is out of bounds.`);
          }
          const table = doc.Tables.Item(tableIndex);
          // WdTableFormatApply flags can be combined (e.g., Borders | Shading | Font | Color | AutoFit | HeadingRows | FirstColumn | LastColumn | LastRow)
          // Example: wdTableFormatApplyBorders = 1, wdTableFormatApplyShading = 2, etc.
          // Default apply flags might vary, check Word documentation. Let's assume applying all is reasonable if not specified.
          const defaultApplyFlags = 1+2+4+8+16+32+64+128+256; // Example: Apply all common flags
          table.AutoFormat(formatName, applyFormatting ?? defaultApplyFlags);
      } catch (error) {
          console.error(`Failed to apply auto format to table ${tableIndex}:`, error);
          throw new Error(`Failed to apply table auto format. Error: ${error}`);
      }
  }

  // --- Image Methods ---

  /**
   * Inserts a picture at the current selection as an inline shape.
   * @param filePath Path to the image file.
   * @param linkToFile Link to the file instead of embedding (optional).
   * @param saveWithDocument Save the image with the document (optional, relevant if linked).
   */
  public async insertPicture(filePath: string, linkToFile: boolean = false, saveWithDocument: boolean = true): Promise<any /* Word.InlineShape */> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      const inlineShape = selection.InlineShapes.AddPicture(filePath, linkToFile, saveWithDocument, selection.Range);
      return inlineShape;
    } catch (error) {
      console.error(`Failed to insert picture from ${filePath}:`, error);
      throw new Error(`Failed to insert picture. Error: ${error}`);
    }
  }

   /**
   * Sets the size of an inline shape (e.g., a picture).
   * Assumes the shape is identified by its index in the active document's InlineShapes collection.
   * @param shapeIndex 1-based index of the inline shape.
   * @param heightPoints Height in points. Use -1 to keep original or maintain aspect ratio if width is set.
   * @param widthPoints Width in points. Use -1 to keep original or maintain aspect ratio if height is set.
   * @param lockAspectRatio Lock aspect ratio when resizing (default: true).
   */
   public async setInlinePictureSize(shapeIndex: number, heightPoints: number, widthPoints: number, lockAspectRatio: boolean = true): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
          if (shapeIndex <= 0 || shapeIndex > doc.InlineShapes.Count) {
              throw new Error(`InlineShape index ${shapeIndex} is out of bounds.`);
          }
          const shape = doc.InlineShapes.Item(shapeIndex);

          // Store original aspect ratio if needed
          const originalHeight = shape.Height;
          const originalWidth = shape.Width;
          // const aspectRatio = originalWidth / originalHeight; // Not needed if relying on LockAspectRatio

          shape.LockAspectRatio = lockAspectRatio ? -1 : 0; // msoTrue = -1, msoFalse = 0

          if (heightPoints > 0 && widthPoints > 0) {
              // Set both, respecting lock aspect ratio if enabled
              if (lockAspectRatio) {
                 // Determine dominant dimension change if aspect ratio is locked
                 const heightRatio = heightPoints / originalHeight;
                 const widthRatio = widthPoints / originalWidth;
                 if (widthRatio > heightRatio) {
                    shape.Width = widthPoints; // Width change is greater, height adjusts
                 } else {
                    shape.Height = heightPoints; // Height change is greater, width adjusts
                 }
              } else {
                 shape.Height = heightPoints;
                 shape.Width = widthPoints;
              }
          } else if (heightPoints > 0) {
              shape.Height = heightPoints; // Width adjusts if aspect ratio locked
          } else if (widthPoints > 0) {
              shape.Width = widthPoints; // Height adjusts if aspect ratio locked
          }
          // If both are <= 0, size remains unchanged

      } catch (error) {
          console.error(`Failed to set size for inline shape ${shapeIndex}:`, error);
          throw new Error(`Failed to set inline picture size. Error: ${error}`);
      }
  }

  // Note: Positioning inline shapes is limited. For more control, convert to a floating Shape.
  // Methods for floating shapes (doc.Shapes) would be needed for complex positioning (Left, Top, Relative anchors).

  // --- Header/Footer Methods ---

  /**
   * Gets a specific header or footer object from a section.
   * @param sectionIndex 1-based index of the section.
   * @param headerFooterType WdHeaderFooterIndex enum value (1=Primary, 2=FirstPage, 3=EvenPages).
   * @param isHeader True for header, false for footer.
   */
  public async getHeaderFooter(sectionIndex: number, headerFooterType: number, isHeader: boolean): Promise<any /* Word.HeaderFooter */> {
      const doc = await this.getActiveDocument();
      try {
          if (sectionIndex <= 0 || sectionIndex > doc.Sections.Count) {
              throw new Error(`Section index ${sectionIndex} is out of bounds.`);
          }
          const section = doc.Sections.Item(sectionIndex);
          const headersFooters = isHeader ? section.Headers : section.Footers;

          // WdHeaderFooterIndex: wdHeaderFooterPrimary = 1, wdHeaderFooterFirstPage = 2, wdHeaderFooterEvenPages = 3
          if (headerFooterType < 1 || headerFooterType > 3) {
             throw new Error(`Invalid header/footer type: ${headerFooterType}. Use 1, 2, or 3.`);
          }

          const headerFooter = headersFooters.Item(headerFooterType);
          if (!headerFooter?.Exists) {
             // Depending on settings (like DifferentFirstPage, DifferentOddAndEvenPages), the requested type might not exist.
             // Handle this gracefully, maybe return null or throw a specific error.
             // For now, let's throw.
             throw new Error(`The requested ${isHeader ? 'header' : 'footer'} type (${headerFooterType}) does not exist or is not active for section ${sectionIndex}. Check document settings.`);
          }
          return headerFooter;
      } catch (error) {
          console.error(`Failed to get ${isHeader ? 'header' : 'footer'} type ${headerFooterType} for section ${sectionIndex}:`, error);
          throw new Error(`Failed to get header/footer. Error: ${error}`);
      }
  }

  /**
   * Sets the text for a specific header or footer. Replaces existing content.
   * @param sectionIndex 1-based index of the section.
   * @param headerFooterType WdHeaderFooterIndex enum value (1=Primary, 2=FirstPage, 3=EvenPages).
   * @param isHeader True for header, false for footer.
   * @param text The text to set.
   */
  public async setHeaderFooterText(sectionIndex: number, headerFooterType: number, isHeader: boolean, text: string): Promise<void> {
      try {
          const headerFooter = await this.getHeaderFooter(sectionIndex, headerFooterType, isHeader);
          headerFooter.Range.Text = text;
      } catch (error) {
          console.error(`Failed to set text for ${isHeader ? 'header' : 'footer'} type ${headerFooterType} section ${sectionIndex}:`, error);
          // Re-throw error as it likely indicates a real issue (invalid index, etc.)
          throw error;
      }
  }

  // --- Page Setup Methods ---

  /**
   * Sets the page margins for the active document.
   * @param topPoints Top margin in points.
   * @param bottomPoints Bottom margin in points.
   * @param leftPoints Left margin in points.
   * @param rightPoints Right margin in points.
   */
  public async setPageMargins(topPoints: number, bottomPoints: number, leftPoints: number, rightPoints: number): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
          const pageSetup = doc.PageSetup;
          pageSetup.TopMargin = topPoints;
          pageSetup.BottomMargin = bottomPoints;
          pageSetup.LeftMargin = leftPoints;
          pageSetup.RightMargin = rightPoints;
      } catch (error) {
          console.error("Failed to set page margins:", error);
          throw new Error(`Failed to set page margins. Error: ${error}`);
      }
  }

  /**
   * Sets the page orientation for the active document.
   * @param orientation WdOrientation enum value (0=Portrait, 1=Landscape).
   */
  public async setPageOrientation(orientation: number): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
          // WdOrientation: wdOrientPortrait = 0, wdOrientLandscape = 1
          doc.PageSetup.Orientation = orientation;
      } catch (error) {
          console.error("Failed to set page orientation:", error);
          throw new Error(`Failed to set page orientation. Error: ${error}`);
      }
  }

  /**
   * Sets the paper size for the active document.
   * @param paperSize WdPaperSize enum value (e.g., 1=Letter, 8=A4).
   */
  public async setPaperSize(paperSize: number): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
          // WdPaperSize enumeration (e.g., wdPaperLetter = 1, wdPaperA4 = 8)
          doc.PageSetup.PaperSize = paperSize;
      } catch (error) {
          console.error("Failed to set paper size:", error);
          throw new Error(`Failed to set paper size. Error: ${error}`);
      }
  }

  // --- Cursor/Selection Methods ---

  /**
   * Moves the cursor to the start of the document.
   */
  public async moveCursorToStart(): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      selection.HomeKey(6); // wdStory = 6
    } catch (error) {
      console.error("Failed to move cursor to start:", error);
      throw new Error(`Failed to move cursor to start. Error: ${error}`);
    }
  }

  /**
   * Moves the cursor to the end of the document.
   */
  public async moveCursorToEnd(): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      selection.EndKey(6); // wdStory = 6
    } catch (error) {
      console.error("Failed to move cursor to end:", error);
      throw new Error(`Failed to move cursor to end. Error: ${error}`);
    }
  }

  /**
   * Moves the cursor by the specified unit and count.
   * @param unit WdUnits enum value (e.g., 1=Character, 2=Word, 3=Sentence, etc.)
   * @param count Number of units to move. Positive moves forward, negative moves backward.
   * @param extend Whether to extend the selection (true) or move the insertion point (false).
   */
  public async moveCursor(unit: number, count: number, extend: boolean = false): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      // WdUnits: wdCharacter = 1, wdWord = 2, wdSentence = 3, wdParagraph = 4, wdLine = 5, wdStory = 6, etc.
      if (extend) {
        selection.MoveRight(unit, count, 1); // 1 = wdExtend
      } else {
        selection.MoveRight(unit, count, 0); // 0 = wdMove
      }
    } catch (error) {
      console.error("Failed to move cursor:", error);
      throw new Error(`Failed to move cursor. Error: ${error}`);
    }
  }

  /**
   * Selects the entire document.
   */
  public async selectAll(): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      selection.WholeStory();
    } catch (error) {
      console.error("Failed to select all:", error);
      throw new Error(`Failed to select all. Error: ${error}`);
    }
  }

  /**
   * Selects a specific paragraph by index.
   * @param paragraphIndex 1-based index of the paragraph to select.
   */
  public async selectParagraph(paragraphIndex: number): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      if (paragraphIndex <= 0 || paragraphIndex > doc.Paragraphs.Count) {
        throw new Error(`Paragraph index ${paragraphIndex} is out of bounds.`);
      }
      const paragraph = doc.Paragraphs.Item(paragraphIndex);
      paragraph.Range.Select();
    } catch (error) {
      console.error(`Failed to select paragraph ${paragraphIndex}:`, error);
      throw new Error(`Failed to select paragraph. Error: ${error}`);
    }
  }

  /**
   * Collapses the current selection to its start or end point.
   * @param toStart If true, collapse to start; if false, collapse to end.
   */
  public async collapseSelection(toStart: boolean = true): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      // WdCollapseDirection: wdCollapseStart = 1, wdCollapseEnd = 0
      selection.Collapse(toStart ? 1 : 0);
    } catch (error) {
      console.error("Failed to collapse selection:", error);
      throw new Error(`Failed to collapse selection. Error: ${error}`);
    }
  }

  /**
   * Gets the current selection text.
   * @returns The text of the current selection.
   */
  public async getSelectionText(): Promise<string> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      return selection.Text;
    } catch (error) {
      console.error("Failed to get selection text:", error);
      throw new Error(`Failed to get selection text. Error: ${error}`);
    }
  }

  /**
   * Gets information about the current selection.
   * @returns Object with selection information.
   */
  public async getSelectionInfo(): Promise<{
    text: string;
    start: number;
    end: number;
    isActive: boolean;
    type: number;
  }> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      return {
        text: selection.Text,
        start: selection.Start,
        end: selection.End,
        isActive: selection.Type !== 0, // wdSelectionNone = 0
        type: selection.Type,
      };
    } catch (error) {
      console.error("Failed to get selection info:", error);
      throw new Error(`Failed to get selection info. Error: ${error}`);
    }
  }

  // --- Add more methods for other Word operations ---

}

// Export a singleton instance
export const wordService = new WordService();
