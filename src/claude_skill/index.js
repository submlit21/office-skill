/**
 * Claude Skill entry point for Office Skill.
 *
 * This module provides the JavaScript interface for Claude Code skill integration.
 * It delegates to the Python backend through the secure bridge script.
 */

import { spawn } from 'child_process';
import { existsSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

/**
 * Office Skill handler for Claude Code.
 */
export class OfficeSkill {
  /**
   * Create a new Office Skill instance.
   */
  constructor() {
    this.pythonPath = process.env.PYTHON_PATH || 'python3';
    this.bridgeScript = path.join(__dirname, 'office_bridge.py');
  }

  /**
   * Execute a command via the bridge script.
   * @param {string} command - Command name
   * @param {Object} params - Command parameters
   * @returns {Promise<Object>} Command result
   */
  async _executeBridge(command, params) {
    return new Promise((resolve, reject) => {
      const proc = spawn(this.pythonPath, [this.bridgeScript], {
        stdio: ['pipe', 'pipe', 'pipe']
      });

      let stdout = '';
      let stderr = '';

      proc.stdout.on('data', (data) => {
        stdout += data.toString();
      });

      proc.stderr.on('data', (data) => {
        stderr += data.toString();
      });

      // Send JSON request via stdin
      const request = JSON.stringify({ command, params });
      proc.stdin.write(request);
      proc.stdin.end();

      proc.on('close', (code) => {
        if (code === 0) {
          try {
            const result = JSON.parse(stdout);
            resolve(result);
          } catch (e) {
            reject(new Error(`Failed to parse bridge output: ${e.message}\nOutput: ${stdout}\nStderr: ${stderr}`));
          }
        } else {
          reject(new Error(`Bridge command failed with code ${code}: ${stderr}`));
        }
      });

      proc.on('error', reject);
    });
  }

  /**
   * Validate file path and check existence if required.
   * @param {string} filePath - File path to validate
   * @param {boolean} checkExists - Whether to check file existence
   * @returns {string} Validated file path
   */
  _validateFilePath(filePath, checkExists = true) {
    if (!filePath || typeof filePath !== 'string') {
      throw new Error(`Invalid file path: ${filePath}`);
    }

    // Basic safety checks
    if (filePath.includes('..') || filePath.includes('\0') || filePath.includes('\n')) {
      throw new Error(`Potentially unsafe file path: ${filePath}`);
    }

    if (checkExists && !existsSync(filePath)) {
      throw new Error(`File not found: ${filePath}`);
    }

    return filePath;
  }

  /**
   * Create a new document.
   * @param {Object} params - Creation parameters
   * @returns {Promise<Object>} Creation result
   */
  async createDocument(params) {
    const { type, output, template } = params;

    if (type !== 'docx' && type !== 'xlsx' && type !== 'pptx') {
      throw new Error(`Unsupported document type: ${type}. Use docx, xlsx, or pptx.`);
    }

    // Validate output path (should not exist yet for creation)
    const outputPath = this._validateFilePath(output, false);

    // Validate template if provided
    let templatePath = null;
    if (template) {
      templatePath = this._validateFilePath(template, true);
    }

    return this._executeBridge('create_document', {
      type,
      output: outputPath,
      template: templatePath
    });
  }

  /**
   * Edit an existing document.
   * @param {Object} params - Edit parameters
   * @returns {Promise<Object>} Edit result
   */
  async editDocument(params) {
    const { input, output, operations } = params;

    // Validate input and output paths
    const inputPath = this._validateFilePath(input, true);
    const outputPath = this._validateFilePath(output, false);

    // Parse operations JSON if needed
    let ops;
    try {
      ops = typeof operations === 'string' ? JSON.parse(operations) : operations;
    } catch (e) {
      throw new Error(`Invalid operations JSON: ${e.message}`);
    }

    // Validate operations structure
    if (!Array.isArray(ops)) {
      throw new Error('Operations must be an array');
    }

    // Basic validation of each operation
    for (const op of ops) {
      if (!op || typeof op !== 'object') {
        throw new Error('Each operation must be an object');
      }
      if (!op.action || typeof op.action !== 'string') {
        throw new Error('Each operation must have an "action" string');
      }
    }

    return this._executeBridge('edit_document', {
      input: inputPath,
      output: outputPath,
      operations: ops
    });
  }

  /**
   * Analyze a document.
   * @param {Object} params - Analysis parameters
   * @returns {Promise<Object>} Analysis result
   */
  async analyzeDocument(params) {
    const { input, type } = params;

    // Validate input path
    const inputPath = this._validateFilePath(input, true);

    // If type not provided, infer from extension
    let docType = type;
    if (!docType) {
      const ext = path.extname(inputPath).toLowerCase().slice(1);
      const typeMap = {
        'docx': 'docx', 'doc': 'docx',
        'xlsx': 'xlsx', 'xls': 'xlsx',
        'pptx': 'pptx', 'ppt': 'pptx'
      };
      docType = typeMap[ext] || 'docx';
    }

    if (docType !== 'docx' && docType !== 'xlsx' && docType !== 'pptx') {
      throw new Error(`Unsupported document type: ${docType}. Use docx, xlsx, or pptx.`);
    }

    return this._executeBridge('analyze_document', {
      input: inputPath,
      type: docType
    });
  }

  /**
   * Export a document to another format.
   * @param {Object} params - Export parameters
   * @returns {Promise<Object>} Export result
   */
  async exportDocument(params) {
    const { input, output, format } = params;

    // Validate input and output paths
    const inputPath = this._validateFilePath(input, true);
    const outputPath = this._validateFilePath(output, false);

    // Validate format
    if (!format || typeof format !== 'string') {
      throw new Error('Export format is required');
    }

    return this._executeBridge('export_document', {
      input: inputPath,
      output: outputPath,
      format
    });
  }

  /**
   * Validate spreadsheet formulas.
   * @param {Object} params - Validation parameters
   * @returns {Promise<Object>} Validation result
   */
  async validateSpreadsheet(params) {
    const { input, timeout = 60 } = params;

    // Validate input path
    const inputPath = this._validateFilePath(input, true);

    // Validate timeout
    const timeoutNum = Number(timeout);
    if (isNaN(timeoutNum) || timeoutNum <= 0) {
      throw new Error('Timeout must be a positive number');
    }

    return this._executeBridge('validate_spreadsheet', {
      input: inputPath,
      timeout: timeoutNum
    });
  }
}

// Export default instance
export default new OfficeSkill();