/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

// OpenAI API Configuration
const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';
const LOCAL_STORAGE_API_KEY = 'openai_api_key';

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Set up event handlers
    document.getElementById("generate-text").onclick = generateAndInsertText;
    document.getElementById("save-api-key").onclick = saveApiKey;
    
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
      const apiKeyInput = document.getElementById("api-key") as HTMLInputElement;
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
    const aiResponse = await callOpenAI(apiKey, prompt);
    
    // Insert the response into Word document
    await insertTextIntoDocument(aiResponse);
    
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
async function callOpenAI(apiKey: string, prompt: string): Promise<string> {
  const response = await fetch(OPENAI_API_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'user',
          content: prompt
        }
      ],
      max_tokens: 1000,
      temperature: 0.7
    })
  });
  
  if (!response.ok) {
    const errorData = await response.json();
    throw new Error(errorData.error?.message || `HTTP ${response.status}: ${response.statusText}`);
  }
  
  const data = await response.json();
  return data.choices[0].message.content.trim();
}

/**
 * Insert generated text into the Word document
 */
async function insertTextIntoDocument(text: string): Promise<void> {
  return Word.run(async (context) => {
    // Insert a paragraph at the end of the document
    const paragraph = context.document.body.insertParagraph(text, Word.InsertLocation.end);
    
    // Optional: Style the inserted text
    paragraph.font.color = "#000000";
    paragraph.font.size = 11;
    
    // Add some spacing
    const spacingParagraph = context.document.body.insertParagraph("", Word.InsertLocation.end);
    spacingParagraph.font.size = 6;
    
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