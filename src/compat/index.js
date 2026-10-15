const fs = require('fs');
const path = require('path');
const vm = require('vm');

const { Utilities } = require('./Utilities');
const { PropertiesService } = require('./PropertiesService');
const { ContentService } = require('./ContentService');
const { HtmlService } = require('./HtmlService');
const { LockService } = require('./LockService');
const { CacheService } = require('./CacheService');
const { UrlFetchApp } = require('./UrlFetchApp');
const { SpreadsheetApp } = require('./SpreadsheetApp');

// Register Google Apps Script globals on globalThis
globalThis.Utilities = Utilities;
globalThis.PropertiesService = PropertiesService;
globalThis.ContentService = ContentService;
globalThis.HtmlService = HtmlService;
globalThis.LockService = LockService;
globalThis.CacheService = CacheService;
globalThis.UrlFetchApp = UrlFetchApp;
globalThis.SpreadsheetApp = SpreadsheetApp;

const SCRIPT_FILES = [
  'Code.js',
  'Status.js',
  'Messaging.js',
  'Conversation.js',
  'RoleReminderV2.js',
  'TemplateSetup.js',
  'Tasks.js',
  'Speaking.js',
  'ProgrammePlanning.js',
  'Engine.js',
  'Workflow.js',
  'WebApiV1.js'
];

let loaded = false;

function loadAppsScriptEnvironment(rootDir = path.resolve(__dirname, '../../')) {
  if (loaded) return;

  for (const filename of SCRIPT_FILES) {
    const fullPath = path.join(rootDir, filename);
    if (!fs.existsSync(fullPath)) {
      throw new Error(`Required Apps Script file not found: ${fullPath}`);
    }
    const code = fs.readFileSync(fullPath, 'utf8');
    vm.runInThisContext(code, { filename });
  }

  loaded = true;
}

// Auto-load upon requiring this module
loadAppsScriptEnvironment();

module.exports = {
  loadAppsScriptEnvironment,
  Utilities,
  PropertiesService,
  ContentService,
  HtmlService,
  LockService,
  CacheService,
  UrlFetchApp,
  SpreadsheetApp
};
