/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// OpenAI API Configuration
//const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';
//const LOCAL_STORAGE_API_KEY = 'openai_api_key';

//Apollo API config
const APOLLO_API_URL = 'https://olympai-a782bc8ad30b.herokuapp.com/api';
const LOCAL_STORAGE_API_KEY = 'apollo_api_key';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Set up event handlers
    document.getElementById("generate-text").onclick = generateAndInsertText;
    document.getElementById("summarize-text").onclick = summarizeSelectedText;
    document.getElementById("save-api-key").onclick = saveApiKey;
    document.getElementById("translate-text").onclick = translateSelectedText;

    
    // Load saved API key
    loadSavedApiKey();
  }
});

/**
 * Load saved API key from localStorage
 */
function loadSavedApiKey(): void {
  try {
    const savedApiKey = localStorage.getItem(LOCAL_STORAGE_API_KEY);
    if (savedApiKey) {
      const apiKeyInput = document.getElementById("api-key") as HTMLInputElement; // so that .value etc is possible
      apiKeyInput.value = savedApiKey;
      showStatus("API key loaded from storage", "success");
    }
  } catch (error) {
    console.log("localStorage not available, API key won't be saved");
  }
}

/**
 * Save API key to localStorage
 */
function saveApiKey(): void {
  try {
    const apiKeyInput = document.getElementById("api-key") as HTMLInputElement;
    const apiKey = apiKeyInput.value.trim();
    
    if (!apiKey) {
      showStatus("Please enter an API key", "error");
      return;
    }
    
    localStorage.setItem(LOCAL_STORAGE_API_KEY, apiKey);
    showStatus("API key saved successfully", "success");
  } catch (error) {
    showStatus("Could not save API key (localStorage not available)", "error");
  }
}

/**
 * Generate text using OpenAI API and insert it into the document
 */
export async function generateAndInsertText(): Promise<void> {
  const apiKeyInput = document.getElementById("api-key") as HTMLInputElement;
  const promptInput = document.getElementById("prompt-input") as HTMLTextAreaElement;
  
  const apiKey = apiKeyInput.value.trim();
  const prompt = promptInput.value.trim();
  
  // Validation
  if (!apiKey) {
    showStatus("Please enter your OpenAI API key", "error");
    return;
  }
  
  if (!prompt) {
    showStatus("Please enter a prompt", "error");
    return;
  }
  
  showStatus("Generating text with AI...", "loading");
  
  try {
    // Call OpenAI API
    const aiResponse = await callApollo(apiKey, prompt); //pauses async function to wait for api call function
    
    // Insert the response into Word document
    await insertTextAtCursor(aiResponse);
    
    showStatus("Text generated and inserted successfully!", "success");
    
    // Clear the prompt input
    promptInput.value = "";
    
  } catch (error) {
    console.error("Error:", error);
    showStatus(`Error: ${error.message}`, "error");
  }
}

/**
 * Call OpenAI API to generate text
 */
// async function callOpenAI(apiKey: string, prompt: string): Promise<string> {
//   const response = await fetch(OPENAI_API_URL, {
//     method: 'POST',
//     headers: {
//       'Content-Type': 'application/json',
//       'X-API-KEY': apiKey
//       //'Authorization': `Bearer ${apiKey}`
//     },
//     body: JSON.stringify({
//       model: 'gpt-3.5-turbo',
//       messages: [
//         {
//           role: 'user',
//           content: prompt
//         }
//       ],
//       max_tokens: 1000,
//       temperature: 0.7
//     })
//   });
  
//   if (!response.ok) {
//     const errorData = await response.json();
//     throw new Error(errorData.error?.message || `HTTP ${response.status}: ${response.statusText}`);
//   }
  
//   const data = await response.json();
//   return data.choices[0].message.content.trim();
// }

// Call Apollo API
async function callApollo(apiKey: string, prompt: string): Promise<string> {
  const response = await fetch(APOLLO_API_URL, {
    method: 'POST',
    headers: {
      'X-API-KEY': apiKey,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      prompt: prompt
    })
  });
  
  if (!response.ok) {
    let errorMessage = `HTTP ${response.status}: ${response.statusText}`;
    
    try {
      const errorData = await response.json();
      // Adapt this based on your API's error response format
      errorMessage = errorData.error || errorData.message || errorMessage;
    } catch (parseError) {
      // If error response isn't JSON, use the default message
    }
    
    throw new Error(errorMessage);
  }
  
  const data = await response.json();
  
  // Extract only the text field from the response, ignoring SQL
  if (data.text) {
    return data.text.trim();
  } else {
    throw new Error("No text field found in API response");
  }
}


/**
 * Insert generated text into the Word document
 */
async function insertTextAtCursor(text: string): Promise<void> {
  return Word.run(async (context) => {
    // Text an der aktuellen Cursor-Position einfügen
    context.document.getSelection().insertText(text, Word.InsertLocation.replace);
    
    await context.sync();
  });
}

/**
 * Show status messages to the user
 */
function showStatus(message: string, type: 'success' | 'error' | 'loading'): void {
  const statusElement = document.getElementById("status");
  
  statusElement.textContent = message;
  statusElement.style.display = "block";
  
  // Remove existing status classes
  statusElement.classList.remove("status-success", "status-error", "status-loading");
  
  // Add appropriate class based on type
  switch (type) {
    case 'success':
      statusElement.classList.add("status-success");
      statusElement.style.backgroundColor = "#d4edda";
      statusElement.style.color = "#155724";
      statusElement.style.border = "1px solid #c3e6cb";
      break;
    case 'error':
      statusElement.classList.add("status-error");
      statusElement.style.backgroundColor = "#f8d7da";
      statusElement.style.color = "#721c24";
      statusElement.style.border = "1px solid #f5c6cb";
      break;
    case 'loading':
      statusElement.classList.add("status-loading");
      statusElement.style.backgroundColor = "#d1ecf1";
      statusElement.style.color = "#0c5460";
      statusElement.style.border = "1px solid #b8daff";
      break;
  }
  
  // Auto-hide success and loading messages after 3 seconds
  if (type === 'success' || type === 'loading') {
    setTimeout(() => {
      statusElement.style.display = "none";
    }, 3000);
  }
}

// Summarize function
async function summarizeSelectedText(): Promise<void> {
  const apiKeyInput = document.getElementById("api-key") as HTMLInputElement;
  const apiKey = apiKeyInput.value.trim();

  if (!apiKey) {
    showStatus("Please enter your OpenAI API key", "error");
    return;
  }

  showStatus("Summarizing selected text...", "loading");

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const selectedText = selection.text.trim();

      if (!selectedText) {
        showStatus("No text selected. Please select text to summarize.", "error");
        return;
      }

      const prompt = `Please summarize the following text in a concise paragraph, while keeping the original language:\n\n${selectedText}`;

      const summary = await callApollo(apiKey, prompt);

      // Replace the selected text with the summary
      selection.insertText(summary, Word.InsertLocation.replace);
      await context.sync();

      showStatus("Summary inserted successfully!", "success");
    });
  } catch (error) {
    console.error("Summarization error:", error);
    showStatus(`Error: ${error.message}`, "error");
  }
}

// translate function

async function translateSelectedText(): Promise<void> {
  const apiKeyInput = document.getElementById("api-key") as HTMLInputElement;
  const languageSelect = document.getElementById("target-language") as HTMLSelectElement;
  const toneInput = document.getElementById("tone") as HTMLInputElement;

  const apiKey = apiKeyInput.value.trim();
  const targetLanguage = languageSelect.value;
  const tone = toneInput.value.trim();

  if (!apiKey) {
    showStatus("Please enter your OpenAI API key", "error");
    return;
  }

  showStatus("Translating selected text...", "loading");

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const selectedText = selection.text.trim();
      if (!selectedText) {
        showStatus("No text selected. Please select text to translate.", "error");
        return;
      }

      let prompt = `Translate the following text into ${targetLanguage}`;
      if (tone) prompt += ` with a ${tone} tone`;
      prompt += `:\n\n${selectedText}`;

      const translation = await callApollo(apiKey, prompt);

      // Replace selected text with the translation
      selection.insertText(translation, Word.InsertLocation.replace);
      await context.sync();

      showStatus("Translation inserted successfully!", "success");
    });
  } catch (error) {
    console.error("Translation error:", error);
    showStatus(`Error: ${error.message}`, "error");
  }
}
