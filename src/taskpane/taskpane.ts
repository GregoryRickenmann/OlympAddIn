/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word, localStorage, console, fetch, setTimeout, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement */

// OpenAI API Configuration
//const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';
//const LOCAL_STORAGE_API_KEY = 'openai_api_key';

//Apollo API config
const APOLLO_API_URL = "https://olympai-a782bc8ad30b.herokuapp.com";
const LOCAL_STORAGE_TOKEN_KEY = "apollo_auth_token";
const LOCAL_STORAGE_USER_EMAIL_KEY = "apollo_user_email";

// Authentication state
let authToken: string | null = null;
let userEmail: string | null = null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Initialize authentication state
    initializeAuthState();

    // Set up event handlers
    document.getElementById("login-button").onclick = authenticate;
    document.getElementById("logout-button").onclick = logout;
    document.getElementById("generate-text").onclick = generateAndInsertText;
    document.getElementById("summarize-text").onclick = summarizeSelectedText;
    document.getElementById("translate-text").onclick = translateSelectedText;

    // Set up example card click handlers
    setupExampleCards();
  }
});

// Initialize authentication state on app load
function initializeAuthState(): void {
  try {
    authToken = localStorage.getItem(LOCAL_STORAGE_TOKEN_KEY);
    userEmail = localStorage.getItem(LOCAL_STORAGE_USER_EMAIL_KEY);

    if (authToken && userEmail) {
      showLoggedInState(userEmail);
    } else {
      showLoggedOutState();
    }
  } catch (error) {
    console.error("Error initializing auth state:", error);
    showLoggedOutState();
  }
}

// Authenticate user
async function authenticate(): Promise<void> {
  const emailInput = document.getElementById("auth-email") as HTMLInputElement;
  const passwordInput = document.getElementById("auth-password") as HTMLInputElement;

  if (!emailInput || !emailInput.value.trim()) {
    showStatus("Please enter your email address", "error");
    return;
  }
  if (!passwordInput || !passwordInput.value.trim()) {
    showStatus("Please enter your password", "error");
    return;
  }

  const email = emailInput.value.trim();
  const password = passwordInput.value.trim();

  showStatus("Logging in...", "loading");

  try {
    const response = await fetch(APOLLO_API_URL + "/word-addin-api/auth/login", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        email: email,
        password: password,
      }),
    });

    if (response.status === 401) {
      showStatus("Invalid email or password", "error");
      return;
    }

    if (response.status === 429) {
      showStatus("Too many login attempts. Please try again later", "error");
      return;
    }

    if (!response.ok) {
      let errorMessage = `Login failed (${response.status})`;
      try {
        const errorData = await response.json();
        errorMessage = errorData.error || errorData.message || errorMessage;
      } catch (e) {
        // Use default error message if response isn't JSON
      }
      throw new Error(errorMessage);
    }

    const data = await response.json();

    if (!data.token) {
      throw new Error("No authentication token received");
    }

    // Save authentication data
    authToken = data.token;
    userEmail = email;
    saveAuthData(data.token, email);

    // Update UI to logged-in state
    showLoggedInState(email);

    // Clear password field
    passwordInput.value = "";

    showStatus("Successfully logged in!", "success");
  } catch (error) {
    console.error("Authentication error:", error);
    showStatus(error.message || "Login failed", "error");
  }
}

// Logout user
function logout(): void {
  authToken = null;
  userEmail = null;
  clearAuthData();
  showLoggedOutState();
  showStatus("Logged out successfully", "success");
}

/**
 * Save authentication data to localStorage
 */
function saveAuthData(token: string, email: string): void {
  try {
    localStorage.setItem(LOCAL_STORAGE_TOKEN_KEY, token);
    localStorage.setItem(LOCAL_STORAGE_USER_EMAIL_KEY, email);
  } catch (error) {
    console.error("Error saving auth data:", error);
    showStatus("Could not save authentication data (localStorage not available)", "error");
  }
}

/**
 * Clear authentication data from localStorage
 */
function clearAuthData(): void {
  try {
    localStorage.removeItem(LOCAL_STORAGE_TOKEN_KEY);
    localStorage.removeItem(LOCAL_STORAGE_USER_EMAIL_KEY);
  } catch (error) {
    console.error("Error clearing auth data:", error);
  }
}

/**
 * Show logged-in UI state
 */
function showLoggedInState(email: string): void {
  document.getElementById("auth-section").style.display = "none";
  document.getElementById("user-info-section").style.display = "block";
  document.getElementById("user-email-display").textContent = email;
}

/**
 * Show logged-out UI state
 */
function showLoggedOutState(): void {
  document.getElementById("auth-section").style.display = "block";
  document.getElementById("user-info-section").style.display = "none";
}

/**
 * Setup example card click handlers
 */
function setupExampleCards(): void {
  const exampleCards = document.querySelectorAll(".example-card");
  exampleCards.forEach((card) => {
    card.addEventListener("click", () => {
      const prompt = card.getAttribute("data-prompt");
      if (prompt) {
        const promptInput = document.getElementById("prompt-input") as HTMLTextAreaElement;
        if (promptInput) {
          promptInput.value = prompt;
        }
      }
    });
  });
}

/**
 * Generate text using Apollo API and insert it into the document
 */
export async function generateAndInsertText(): Promise<void> {
  if (!authToken) {
    showStatus("Please log in first", "error");
    return;
  }

  const promptInput = document.getElementById("prompt-input") as HTMLTextAreaElement;
  const prompt = promptInput.value.trim();

  if (!prompt) {
    showStatus("Please enter a prompt", "error");
    return;
  }

  showStatus("Generating text with AI...", "loading");

  try {
    const aiResponse = await callApolloPrompt(prompt);

    // Insert the response into Word document
    await insertTextAtCursor(aiResponse);

    showStatus("Text generated and inserted successfully!", "success");

    // Clear the prompt input
    promptInput.value = "";
  } catch (error) {
    console.error("Error:", error);
    if (error.message.includes("401")) {
      // Token expired or invalid
      logout();
      showStatus("Session expired. Please log in again", "error");
    } else {
      showStatus(`Error: ${error.message}`, "error");
    }
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

// Call Apollo API for prompts
async function callApolloPrompt(prompt: string): Promise<string> {
  if (!authToken) {
    throw new Error("Authentication required");
  }

  const response = await fetch(APOLLO_API_URL + "/word-addin-api/prompt", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${authToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      prompt: prompt,
    }),
  });

  if (response.status === 401) {
    throw new Error("Authentication failed. Please log in again.");
  }

  if (response.status === 429) {
    throw new Error("Rate limit exceeded. Please try again later.");
  }

  if (response.status === 500) {
    throw new Error("Server error. Please try again later.");
  }

  if (!response.ok) {
    let errorMessage = `HTTP ${response.status}: ${response.statusText}`;

    try {
      const errorData = await response.json();
      errorMessage = errorData.error || errorData.message || errorMessage;
    } catch (parseError) {
      // If error response isn't JSON, use the default message
    }

    throw new Error(errorMessage);
  }

  const data = await response.json();

  // Extract only the text field from the response
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
function showStatus(message: string, type: "success" | "error" | "loading"): void {
  const statusElement = document.getElementById("status");

  statusElement.textContent = message;
  statusElement.style.display = "block";

  // Remove existing status classes
  statusElement.classList.remove("status-success", "status-error", "status-loading");

  // Add appropriate class based on type
  switch (type) {
    case "success":
      statusElement.classList.add("status-success");
      statusElement.style.backgroundColor = "#d4edda";
      statusElement.style.color = "#155724";
      statusElement.style.border = "1px solid #c3e6cb";
      break;
    case "error":
      statusElement.classList.add("status-error");
      statusElement.style.backgroundColor = "#f8d7da";
      statusElement.style.color = "#721c24";
      statusElement.style.border = "1px solid #f5c6cb";
      break;
    case "loading":
      statusElement.classList.add("status-loading");
      statusElement.style.backgroundColor = "#d1ecf1";
      statusElement.style.color = "#0c5460";
      statusElement.style.border = "1px solid #b8daff";
      break;
  }

  // Auto-hide success and loading messages after 3 seconds
  if (type === "success" || type === "loading") {
    setTimeout(() => {
      statusElement.style.display = "none";
    }, 3000);
  }
}

// Summarize function
async function summarizeSelectedText(): Promise<void> {
  if (!authToken) {
    showStatus("Please log in first", "error");
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

      const summary = await callApolloPrompt(prompt);

      // Replace the selected text with the summary
      selection.insertText(summary, Word.InsertLocation.replace);
      await context.sync();

      showStatus("Summary inserted successfully!", "success");
    });
  } catch (error) {
    console.error("Summarization error:", error);
    if (error.message.includes("401")) {
      logout();
      showStatus("Session expired. Please log in again", "error");
    } else {
      showStatus(`Error: ${error.message}`, "error");
    }
  }
}

// translate function

async function translateSelectedText(): Promise<void> {
  if (!authToken) {
    showStatus("Please log in first", "error");
    return;
  }

  const languageSelect = document.getElementById("target-language") as HTMLSelectElement;
  const toneInput = document.getElementById("tone") as HTMLInputElement;

  const targetLanguage = languageSelect.value;
  const tone = toneInput.value.trim();

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

      const translation = await callApolloPrompt(prompt);

      // Replace selected text with the translation
      selection.insertText(translation, Word.InsertLocation.replace);
      await context.sync();

      showStatus("Translation inserted successfully!", "success");
    });
  } catch (error) {
    console.error("Translation error:", error);
    if (error.message.includes("401")) {
      logout();
      showStatus("Session expired. Please log in again", "error");
    } else {
      showStatus(`Error: ${error.message}`, "error");
    }
  }
}
