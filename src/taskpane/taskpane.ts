/*
 * Copyright (c) Brian Dawley and Shumaker & Sieffert, P.A. All rights reserved.
 */

/* global document, Office, Word */

// This interface holds the data we load from Word
interface PromptData {
  buttonText: string;
  textToCopy: string;
}

// This interface holds UI items (labels or buttons)
interface UIItem {
  type: "label" | "button";
  text: string;
  data?: PromptData; // Only for buttons
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("scan-button").onclick = scanAndBuildUI;
    scanAndBuildUI(); // Run the scan immediately on load
  }
});

/**
 * Scans the document for all headings and prompt text,
 * then builds the UI in one consolidated operation.
 */
async function scanAndBuildUI() {
  const statusMessage = document.getElementById("status-message");
  const container = document.getElementById("button-container");
  container.innerHTML = ""; // Clear old UI
  statusMessage.innerText = "Scanning document...";

  try {
    await Word.run(async (context) => {
      // 1. Load all paragraphs with their style and text
      const paragraphs = context.document.body.paragraphs.load("items/styleBuiltIn, items/text");
      await context.sync();

      // 2. Find all headings
      const allHeadings = [];
      for (let i = 0; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        if (p.styleBuiltIn.startsWith("Heading")) {
          allHeadings.push({
            paragraph: p,
            text: p.text,
            index: i,
          });
        }
      }

      if (allHeadings.length === 0) {
        statusMessage.innerText = "No headings found in this document.";
        return;
      }

      // 3. Prepare to load the text for all "button" headings
      const uiItemsToCreate: UIItem[] = [];
      const rangesToLoad = []; // We will load all text ranges at once

      for (let i = 0; i < allHeadings.length; i++) {
        const currentHeading = allHeadings[i];
        const headingText = currentHeading.text.trim();

        if (headingText.startsWith("*")) {
          // This is a button. Find its content.
          const nextHeading = allHeadings[i + 1];

          // Get the range *after* this heading
          const startRange = currentHeading.paragraph.getRange("End");
          
          // Get the range *before* the next heading (or end of doc)
          const endRange = nextHeading
            ? nextHeading.paragraph.getRange("Start")
            : context.document.body.getRange("End");

          // Get the content range and queue it for loading
          const contentRange = startRange.expandTo(endRange);
          contentRange.load("text");

          // Prepare the data structure
          const promptData = {
            buttonText: headingText.substring(1).trim(),
            textToCopy: "", // Will be filled in after the sync
            _range: contentRange, // Internal temp reference
          };

          uiItemsToCreate.push({
            type: "button",
            text: promptData.buttonText,
            data: promptData,
          });
          rangesToLoad.push(promptData);

        } else if (headingText.startsWith("_")) {
          // This is a label
          uiItemsToCreate.push({
            type: "label",
            text: headingText.substring(1).trim(),
          });
        }
      }

      // 4. Run the second sync to load all the text for the buttons
      await context.sync();

      // 5. Now that sync is done, retrieve the text from the loaded ranges
      for (const prompt of rangesToLoad) {
        prompt.textToCopy = prompt._range.text.trim();
      }

      // 6. Build the final UI
      statusMessage.innerText = ""; // Clear "Scanning..."
      if (uiItemsToCreate.length === 0) {
          statusMessage.innerText = "No prompts found (use '*' or '_' in headings).";
          return;
      }

      uiItemsToCreate.forEach((item) => {
        if (item.type === "label") {
          const label = document.createElement("h3");
          label.className = "ms-font-l";
          label.innerText = item.text;
          container.appendChild(label);
        } else if (item.type === "button") {
          const button = document.createElement("button");
          button.className = "ms-Button";
          button.innerText = item.text;
          
          // Attach the NEW, simple click handler
          button.onclick = () => {
            copyTextToClipboard(item.data.textToCopy, item.data.buttonText);
          };
          container.appendChild(button);
        }
      });
    });
  } catch (error) {
    console.error("Error in scanAndBuildUI:", error);
    statusMessage.innerText = "Error scanning document. See console.";
  }
}

/**
 * Copies a given string to the clipboard.
 * This function no longer uses Word.run() and is much safer.
 */
async function copyTextToClipboard(textToCopy: string, buttonText: string) {
  const statusMessage = document.getElementById("status-message");
  
  // DEBUGGING: Log 1
  console.log(`Button clicked for: "${buttonText}"`);
  statusMessage.innerText = `Copying "${buttonText}"...`;

  try {
    if (textToCopy.length === 0) {
      console.warn("Text to copy is empty. (Was it empty in the doc?)");
      statusMessage.innerText = "Warning: No text found to copy.";
      return;
    }
    
    // DEBUGGING: Log 2
    console.log(`Text to copy (length ${textToCopy.length}): "${textToCopy}"`);

    // --- CRITICAL PART ---
    await navigator.clipboard.writeText(textToCopy);
    
    // DEBUGGING: Log 3
    console.log("Successfully copied to clipboard.");
    statusMessage.innerText = `Copied "${buttonText}" to clipboard!`;

  } catch (error) {
    // DEBUGGING: Log 4 - This will now only catch clipboard errors
    console.error("Error copying to clipboard:", error);
    statusMessage.innerText = "Error! See console for details.";
  }
}