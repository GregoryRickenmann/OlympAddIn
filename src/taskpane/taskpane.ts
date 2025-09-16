/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word, localStorage, console, fetch, setTimeout, HTMLInputElement, HTMLSelectElement, HTMLTextAreaElement */

import MarkdownIt from "markdown-it";

// OpenAI API Configuration
//const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';
//const LOCAL_STORAGE_API_KEY = 'openai_api_key';

//Apollo API config
// const APOLLO_API_URL = "https://olympai-a782bc8ad30b.herokuapp.com";
const APOLLO_API_URL = "https://app.olymp.finance";
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
      } catch {
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

  showStatus("Starting AI text generation...", "loading");
  hideSources();

  try {
    const aiResponse = await callApolloPrompt(prompt);

    // Insert the response into Word document
    await insertTextAtCursor(aiResponse.text);

    // Display sources
    displaySources(aiResponse.sources);

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

// Call Apollo API for prompts using async task structure
async function callApolloPrompt(prompt: string): Promise<{ text: string; sources: any }> {
  if (!authToken) {
    throw new Error("Authentication required");
  }

  // Step 1: Start the async task
  const taskId = await startAsyncQuery(prompt);

  // Step 2: Poll for results
  return await pollTaskStatus(taskId);
}

// Start an async query and return the task ID
async function startAsyncQuery(prompt: string): Promise<string> {
  const response = await fetch(APOLLO_API_URL + "/word-addin-api/prompt/start", {
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
    } catch {
      // If error response isn't JSON, use the default message
    }

    throw new Error(errorMessage);
  }

  const data = await response.json();

  if (!data.task_id) {
    throw new Error("No task ID received from server");
  }

  return data.task_id;
}

// Poll task status until completion
async function pollTaskStatus(taskId: string): Promise<{ text: string; sources: any }> {
  const maxAttempts = 120; // Maximum 2 minutes (120 * 1s)
  let attempts = 0;

  while (attempts < maxAttempts) {
    try {
      const response = await fetch(APOLLO_API_URL + `/word-addin-api/prompt/status/${taskId}`, {
        method: "GET",
        headers: {
          Authorization: `Bearer ${authToken}`,
        },
      });

      if (response.status === 401) {
        throw new Error("Authentication failed. Please log in again.");
      }

      if (!response.ok) {
        let errorMessage = `HTTP ${response.status}: ${response.statusText}`;
        try {
          const errorData = await response.json();
          errorMessage = errorData.error || errorData.message || errorMessage;
        } catch {
          // Use default error message
        }
        throw new Error(errorMessage);
      }

      const data = await response.json();

      // Handle different task states
      switch (data.status) {
        case "success":
          // Task completed successfully
          if (!data.result || !data.result.text) {
            throw new Error("No text field found in API response");
          }
          return {
            text: data.result.text.trim(),
            sources: data.result.sources || {},
          };

        case "failure":
          // Task failed
          throw new Error(data.error || "Query processing failed");

        case "pending":
        case "started":
          // Task still running, update status message
          if (data.progress) {
            updateStatusMessage(data.progress);
          }
          break;

        default:
          // Unknown status, continue polling
          break;
      }

      // Wait 1 second before next poll
      await new Promise((resolve) => setTimeout(resolve, 1000));
      attempts++;
    } catch (error) {
      // If it's an authentication error, re-throw it
      if (error.message.includes("401") || error.message.includes("Authentication")) {
        throw error;
      }
      // For other errors, wait and retry
      await new Promise((resolve) => setTimeout(resolve, 1000));
      attempts++;
    }
  }

  // Timeout reached
  throw new Error("Query processing timeout. Please try again.");
}

// Update status message during polling
function updateStatusMessage(progress: string): void {
  showStatus(progress, "loading");
}

/**
 * Insert generated text into the Word document with markdown formatting
 */
async function insertTextAtCursor(text: string): Promise<void> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    
    // Initialize markdown parser
    const md = new MarkdownIt({
      html: false,
      breaks: true,
      linkify: true,
      typographer: true
    });
    
    // Parse markdown to tokens
    const tokens = md.parse(text, {});
    
    // Clear selection and get the range for insertion
    selection.clear();
    
    // Process each token and apply appropriate formatting
    // We'll track the current insertion point and update it after each element
    await processMarkdownTokens(context, selection, tokens);
    
    await context.sync();
  });
}

/**
 * Process markdown tokens and apply Word formatting
 */
async function processMarkdownTokens(context: Word.RequestContext, selection: Word.Range, tokens: any[]): Promise<void> {
  // Track the current insertion point
  let currentRange = selection;
  
  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];
    
    switch (token.type) {
      case 'heading_open':
        currentRange = await insertHeading(context, currentRange, token, tokens[i + 1]);
        i++; // Skip content token
        break;
        
      case 'paragraph_open':
        const paraContent = tokens[i + 1];
        if (paraContent && paraContent.type === 'inline') {
          currentRange = await insertFormattedParagraph(context, currentRange, paraContent);
        }
        i++; // Skip content and close tokens
        break;
        
      case 'bullet_list_open':
        currentRange = await insertBulletList(context, currentRange, tokens, i);
        // Skip to end of list
        let listDepth = 1;
        i++;
        while (i < tokens.length && listDepth > 0) {
          if (tokens[i].type === 'bullet_list_open') listDepth++;
          if (tokens[i].type === 'bullet_list_close') listDepth--;
          i++;
        }
        i--; // Adjust for loop increment
        break;
        
      case 'ordered_list_open':
        currentRange = await insertOrderedList(context, currentRange, tokens, i);
        // Skip to end of list
        let orderedDepth = 1;
        i++;
        while (i < tokens.length && orderedDepth > 0) {
          if (tokens[i].type === 'ordered_list_open') orderedDepth++;
          if (tokens[i].type === 'ordered_list_close') orderedDepth--;
          i++;
        }
        i--; // Adjust for loop increment
        break;
        
      case 'blockquote_open':
        currentRange = await insertBlockquote(context, currentRange, tokens, i);
        // Skip to end of blockquote
        let quoteDepth = 1;
        i++;
        while (i < tokens.length && quoteDepth > 0) {
          if (tokens[i].type === 'blockquote_open') quoteDepth++;
          if (tokens[i].type === 'blockquote_close') quoteDepth--;
          i++;
        }
        i--; // Adjust for loop increment
        break;
        
      case 'code_block':
        currentRange = await insertCodeBlock(context, currentRange, token);
        break;
        
      case 'hr':
        currentRange = await insertHorizontalRule(context, currentRange);
        break;
        
      case 'table_open':
        currentRange = await insertTable(context, currentRange, tokens, i);
        // Skip to end of table
        let tableDepth = 1;
        i++;
        while (i < tokens.length && tableDepth > 0) {
          if (tokens[i].type === 'table_open') tableDepth++;
          if (tokens[i].type === 'table_close') tableDepth--;
          i++;
        }
        i--; // Adjust for loop increment
        break;
        
      default:
        // Handle other tokens as plain text
        if (token.content) {
          const textRange = currentRange.insertText(token.content + '\n', Word.InsertLocation.after);
          currentRange = textRange;
        }
        break;
    }
  }
}

/**
 * Insert formatted heading
 */
async function insertHeading(context: Word.RequestContext, selection: Word.Range, openToken: any, contentToken: any): Promise<Word.Range> {
  if (!contentToken || !contentToken.content) return selection;
  
  const level = parseInt(openToken.tag.charAt(1)); // Extract number from h1, h2, etc.
  const headingText = contentToken.content;
  
  const paragraph = selection.insertParagraph(headingText, Word.InsertLocation.after);
  
  // Apply heading style based on level
  switch (level) {
    case 1:
      paragraph.styleBuiltIn = Word.BuiltInStyleName.heading1;
      break;
    case 2:
      paragraph.styleBuiltIn = Word.BuiltInStyleName.heading2;
      break;
    case 3:
      paragraph.styleBuiltIn = Word.BuiltInStyleName.heading3;
      break;
    default:
      paragraph.styleBuiltIn = Word.BuiltInStyleName.heading4;
      break;
  }
  
  return paragraph.getRange(Word.RangeLocation.end);
}

/**
 * Insert formatted paragraph with inline formatting
 */
async function insertFormattedParagraph(context: Word.RequestContext, selection: Word.Range, inlineToken: any): Promise<Word.Range> {
  if (!inlineToken.children || inlineToken.children.length === 0) {
    const emptyParagraph = selection.insertParagraph('', Word.InsertLocation.after);
    return emptyParagraph.getRange(Word.RangeLocation.end);
  }
  
  const paragraph = selection.insertParagraph('', Word.InsertLocation.after);
  
  // Track formatting state
  let isBold = false;
  let isItalic = false;
  let linkHref = '';
  let pendingText = '';
  
  const applyPendingText = () => {
    if (pendingText) {
      const range = paragraph.insertText(pendingText, Word.InsertLocation.end);
      if (isBold) range.font.bold = true;
      if (isItalic) range.font.italic = true;
      if (linkHref) {
        // Add hyperlink formatting
        range.font.color = '#0066cc';
        range.font.underline = Word.UnderlineType.single;
      }
      pendingText = '';
    }
  };
  
  for (let i = 0; i < inlineToken.children.length; i++) {
    const child = inlineToken.children[i];
    
    switch (child.type) {
      case 'text':
        pendingText += child.content;
        break;
        
      case 'strong_open':
        applyPendingText();
        isBold = true;
        break;
        
      case 'strong_close':
        applyPendingText();
        isBold = false;
        break;
        
      case 'em_open':
        applyPendingText();
        isItalic = true;
        break;
        
      case 'em_close':
        applyPendingText();
        isItalic = false;
        break;
        
      case 'code_inline':
        applyPendingText();
        const codeRange = paragraph.insertText(child.content, Word.InsertLocation.end);
        codeRange.font.name = 'Courier New';
        codeRange.font.color = '#d63384';
        codeRange.font.highlightColor = '#f8f9fa';
        break;
        
      case 'link_open':
        applyPendingText();
        linkHref = child.attrs?.find(attr => attr[0] === 'href')?.[1] || '';
        break;
        
      case 'link_close':
        applyPendingText();
        linkHref = '';
        break;
        
      default:
        if (child.content) {
          pendingText += child.content;
        }
        break;
    }
  }
  
  // Apply any remaining text
  applyPendingText();
  
  return paragraph.getRange(Word.RangeLocation.end);
}

/**
 * Insert bullet list
 */
async function insertBulletList(context: Word.RequestContext, selection: Word.Range, tokens: any[], startIndex: number): Promise<Word.Range> {
  let i = startIndex + 1;
  let lastParagraph: Word.Paragraph | null = null;
  
  while (i < tokens.length && tokens[i].type !== 'bullet_list_close') {
    if (tokens[i].type === 'list_item_open' && tokens[i + 1] && tokens[i + 1].type === 'paragraph_open') {
      const contentToken = tokens[i + 2];
      if (contentToken && contentToken.type === 'inline' && contentToken.content) {
        lastParagraph = selection.insertParagraph('• ' + contentToken.content, Word.InsertLocation.after);
        lastParagraph.leftIndent = 18;
        selection = lastParagraph.getRange(Word.RangeLocation.end);
      }
      i += 3; // Skip item_open, para_open, inline, para_close, item_close
    } else {
      i++;
    }
  }
  
  return lastParagraph ? lastParagraph.getRange(Word.RangeLocation.end) : selection;
}

/**
 * Insert ordered list
 */
async function insertOrderedList(context: Word.RequestContext, selection: Word.Range, tokens: any[], startIndex: number): Promise<Word.Range> {
  let i = startIndex + 1;
  let itemNumber = 1;
  let lastParagraph: Word.Paragraph | null = null;
  
  while (i < tokens.length && tokens[i].type !== 'ordered_list_close') {
    if (tokens[i].type === 'list_item_open' && tokens[i + 1] && tokens[i + 1].type === 'paragraph_open') {
      const contentToken = tokens[i + 2];
      if (contentToken && contentToken.type === 'inline' && contentToken.content) {
        lastParagraph = selection.insertParagraph(`${itemNumber}. ${contentToken.content}`, Word.InsertLocation.after);
        lastParagraph.leftIndent = 18;
        selection = lastParagraph.getRange(Word.RangeLocation.end);
        itemNumber++;
      }
      i += 3; // Skip item_open, para_open, inline, para_close, item_close
    } else {
      i++;
    }
  }
  
  return lastParagraph ? lastParagraph.getRange(Word.RangeLocation.end) : selection;
}

/**
 * Insert blockquote
 */
async function insertBlockquote(context: Word.RequestContext, selection: Word.Range, tokens: any[], startIndex: number): Promise<Word.Range> {
  let i = startIndex + 1;
  let lastParagraph: Word.Paragraph | null = null;
  
  while (i < tokens.length && tokens[i].type !== 'blockquote_close') {
    if (tokens[i].type === 'paragraph_open' && tokens[i + 1] && tokens[i + 1].type === 'inline') {
      const contentToken = tokens[i + 1];
      if (contentToken.content) {
        lastParagraph = selection.insertParagraph(contentToken.content, Word.InsertLocation.after);
        lastParagraph.styleBuiltIn = Word.BuiltInStyleName.quote;
        lastParagraph.leftIndent = 36; // Add left indent for quote styling
        lastParagraph.font.italic = true;
        lastParagraph.font.color = '#6c757d';
        selection = lastParagraph.getRange(Word.RangeLocation.end);
      }
      i += 2; // Skip para_open, inline, para_close
    } else {
      i++;
    }
  }
  
  return lastParagraph ? lastParagraph.getRange(Word.RangeLocation.end) : selection;
}

/**
 * Insert code block
 */
async function insertCodeBlock(context: Word.RequestContext, selection: Word.Range, token: any): Promise<Word.Range> {
  if (!token.content) return selection;
  
  const codeBlock = selection.insertParagraph(token.content, Word.InsertLocation.after);
  codeBlock.font.name = 'Courier New';
  codeBlock.font.size = 10;
  codeBlock.font.color = '#212529';
  codeBlock.font.highlightColor = '#f8f9fa';
  codeBlock.leftIndent = 36;
  codeBlock.spaceAfter = 6;
  codeBlock.spaceBefore = 6;
  
  return codeBlock.getRange(Word.RangeLocation.end);
}

/**
 * Insert horizontal rule
 */
async function insertHorizontalRule(context: Word.RequestContext, selection: Word.Range): Promise<Word.Range> {
  const hrParagraph = selection.insertParagraph('─────────────────────────────────────────', Word.InsertLocation.after);
  hrParagraph.alignment = Word.Alignment.centered;
  hrParagraph.font.color = '#dee2e6';
  hrParagraph.spaceAfter = 12;
  hrParagraph.spaceBefore = 12;
  
  return hrParagraph.getRange(Word.RangeLocation.end);
}

/**
 * Insert table from markdown tokens
 */
async function insertTable(context: Word.RequestContext, selection: Word.Range, tokens: any[], startIndex: number): Promise<Word.Range> {
  // Parse table structure
  const tableData: string[][] = [];
  let headerRow: string[] = [];
  let i = startIndex + 1;

  // Find table header
  while (i < tokens.length && tokens[i].type !== 'table_close') {
    if (tokens[i].type === 'thead_open') {
      // Process header row
      while (i < tokens.length && tokens[i].type !== 'thead_close') {
        if (tokens[i].type === 'tr_open') {
          headerRow = [];
          i++;
          while (i < tokens.length && tokens[i].type !== 'tr_close') {
            if (tokens[i].type === 'th_open' && tokens[i + 1] && tokens[i + 1].type === 'inline') {
              headerRow.push(tokens[i + 1].content || '');
              i += 2; // Skip th_open, inline, th_close
            } else {
              i++;
            }
          }
          if (headerRow.length > 0) {
            tableData.push(headerRow);
          }
        }
        i++;
      }
    } else if (tokens[i].type === 'tbody_open') {
      // Process body rows
      while (i < tokens.length && tokens[i].type !== 'tbody_close') {
        if (tokens[i].type === 'tr_open') {
          const row: string[] = [];
          i++;
          while (i < tokens.length && tokens[i].type !== 'tr_close') {
            if (tokens[i].type === 'td_open' && tokens[i + 1] && tokens[i + 1].type === 'inline') {
              row.push(tokens[i + 1].content || '');
              i += 2; // Skip td_open, inline, td_close
            } else {
              i++;
            }
          }
          if (row.length > 0) {
            tableData.push(row);
          }
        }
        i++;
      }
    } else {
      i++;
    }
  }

  // Create Word table if we have data
  if (tableData.length > 0 && tableData[0].length > 0) {
    const rowCount = tableData.length;
    const colCount = tableData[0].length;
    
    // Insert empty paragraph for table placement
    const tableParagraph = selection.insertParagraph('', Word.InsertLocation.after);
    
    // Create table
    const table = tableParagraph.insertTable(rowCount, colCount, Word.InsertLocation.after);
    
    // Populate table data
    for (let row = 0; row < tableData.length; row++) {
      for (let col = 0; col < tableData[row].length && col < colCount; col++) {
        const cell = table.getCell(row, col);
        cell.value = tableData[row][col] || '';
        
        // Style header row
        if (row === 0) {
          cell.body.font.bold = true;
          cell.body.font.color = '#2c3e50';
          cell.horizontalAlignment = Word.Alignment.centered;
        }
      }
    }
    
    // Apply table styling
    table.styleBuiltIn = Word.BuiltInStyleName.gridTable4_Accent1;
    table.horizontalAlignment = Word.Alignment.centered;
    
    // Add spacing around table
    const afterParagraph = selection.insertParagraph('', Word.InsertLocation.after);
    afterParagraph.spaceAfter = 12;
    tableParagraph.spaceBefore = 12;
    
    return afterParagraph.getRange(Word.RangeLocation.end);
  }
  
  return selection;
}

/**
 * Display sources in the taskpane
 */
function displaySources(sources: any): void {
  const sourcesSection = document.getElementById("sources-section");
  const sourcesList = document.getElementById("sources-list");

  if (!sources || Object.keys(sources).length === 0) {
    sourcesSection.style.display = "none";
    return;
  }

  // Clear existing content
  sourcesList.innerHTML = "";

  // Create list items for each source
  for (const [key, value] of Object.entries(sources)) {
    const sourceItem = document.createElement("div");
    sourceItem.className = "source-item";

    const sourceKey = document.createElement("div");
    sourceKey.className = "source-key";
    sourceKey.textContent = getSourceKeyDisplay(key);

    const sourceContent = document.createElement("div");
    sourceContent.className = "source-content";

    // Handle different value types
    if (Array.isArray(value)) {
      sourceContent.appendChild(createArrayDisplay(value, key));
    } else if (typeof value === "object" && value !== null) {
      sourceContent.appendChild(createObjectDisplay(value));
    } else {
      const valueSpan = document.createElement("span");
      valueSpan.className = "source-value-simple";
      valueSpan.textContent = String(value);
      sourceContent.appendChild(valueSpan);
    }

    sourceItem.appendChild(sourceKey);
    sourceItem.appendChild(sourceContent);
    sourcesList.appendChild(sourceItem);
  }

  // Show the sources section
  sourcesSection.style.display = "block";
}

/**
 * Get display name for source key
 */
function getSourceKeyDisplay(key: string): string {
  const keyMappings: Record<string, string> = {
    database: "🗄️ Database",
    web_search: "🔍 Web Search",
    api_connection: "🔗 API Connection",
    files: "📁 Files",
  };
  return keyMappings[key] || key;
}

/**
 * Create display for array values
 */
function createArrayDisplay(array: any[], sourceType: string): HTMLElement {
  const container = document.createElement("div");
  container.className = "source-array";

  array.forEach((item) => {
    const itemElement = document.createElement("div");
    itemElement.className = "source-array-item";

    if (sourceType === "web_search" && typeof item === "object" && item.title && item.url) {
      // Handle web search results
      const link = document.createElement("a");
      link.href = item.url;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.className = "source-link";
      link.textContent = item.title;

      const urlSpan = document.createElement("span");
      urlSpan.className = "source-url";
      urlSpan.textContent = item.url;

      itemElement.appendChild(link);
      itemElement.appendChild(urlSpan);
    } else if (sourceType === "files") {
      // Handle file lists
      const fileIcon = getFileIcon(String(item));
      const fileSpan = document.createElement("span");
      fileSpan.className = "source-file";
      fileSpan.textContent = `${fileIcon} ${item}`;
      itemElement.appendChild(fileSpan);
    } else {
      // Handle other array items
      const valueSpan = document.createElement("span");
      valueSpan.className = "source-value-simple";
      valueSpan.textContent = String(item);
      itemElement.appendChild(valueSpan);
    }

    container.appendChild(itemElement);
  });

  return container;
}

/**
 * Create display for object values
 */
function createObjectDisplay(obj: any): HTMLElement {
  const container = document.createElement("div");
  container.className = "source-object";

  for (const [key, value] of Object.entries(obj)) {
    const itemElement = document.createElement("div");
    itemElement.className = "source-object-item";

    const keySpan = document.createElement("span");
    keySpan.className = "source-object-key";
    keySpan.textContent = key + ":";

    const valueSpan = document.createElement("span");
    valueSpan.className = "source-object-value";
    valueSpan.textContent = String(value);

    itemElement.appendChild(keySpan);
    itemElement.appendChild(valueSpan);
    container.appendChild(itemElement);
  }

  return container;
}

/**
 * Get appropriate icon for file type
 */
function getFileIcon(filename: string): string {
  const extension = filename.split(".").pop()?.toLowerCase() || "";

  const iconMappings: Record<string, string> = {
    pdf: "📄",
    doc: "📝",
    docx: "📝",
    xls: "📊",
    xlsx: "📊",
    ppt: "📽️",
    pptx: "📽️",
    txt: "📄",
    csv: "📊",
    json: "🔧",
    xml: "🔧",
  };

  return iconMappings[extension] || "📄";
}

/**
 * Hide sources section
 */
function hideSources(): void {
  const sourcesSection = document.getElementById("sources-section");
  sourcesSection.style.display = "none";
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

  // Auto-hide success messages after 3 seconds, loading messages stay visible
  if (type === "success") {
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

  showStatus("Starting text summarization...", "loading");
  hideSources();

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
      selection.insertText(summary.text, Word.InsertLocation.replace);
      await context.sync();

      // Display sources
      displaySources(summary.sources);

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

  showStatus("Starting text translation...", "loading");
  hideSources();

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
      selection.insertText(translation.text, Word.InsertLocation.replace);
      await context.sync();

      // Display sources
      displaySources(translation.sources);

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
