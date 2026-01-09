/**
 * Gateway Calibration Tool
 * Listens to MQTT RSSI data, records at specific distances for 1 minute,
 * averages readings, and writes to Excel file
 * Adapted for mosquitto-client message format
 */

const mqtt = require('mqtt');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const brokerUrl = process.env.MQTT_BROKER_URL || 'mqtt://localhost:1883';
const RECORDING_DURATION = 60 * 1000; // 1 minute in milliseconds

class GatewayCalibrationTool {
  constructor() {
    this.client = null;
    this.rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
    this.recordings = [];
    this.isRecording = false;
    this.recordingStartTime = null;
    this.currentGatewayMac = null;
    this.currentDistance = null;
    this.outputFile = null;
  }

  async connect() {
    return new Promise((resolve, reject) => {
      this.client = mqtt.connect(brokerUrl, {
        clientId: `calibration-tool-${Date.now()}`
      });

      this.client.on('connect', () => {
        console.log(`✓ Connected to MQTT broker: ${brokerUrl}\n`);
        resolve();
      });

      this.client.on('error', (error) => {
        console.error('MQTT error:', error);
        reject(error);
      });

      this.client.on('message', (topic, message) => {
        this.handleMessage(topic, message);
      });
    });
  }

  handleMessage(topic, message) {
    if (!this.isRecording) return;

    try {
      const payload = JSON.parse(message.toString());
      
      // Parse mosquitto-client message format:
      // { device_info: {mac: "..."}, data: [{mac: "...", rssi: ...}] }
      // device_info.mac is the GATEWAY MAC
      // data array contains tags/devices detected by that gateway
      if (!payload.device_info || !payload.device_info.mac) return;
      if (!Array.isArray(payload.data)) return;

      const gatewayMac = payload.device_info.mac.toUpperCase();
      
      // Only process messages from the target gateway
      if (gatewayMac !== this.currentGatewayMac) return;

      // Collect RSSI values from all tags detected by this gateway
      payload.data.forEach(item => {
        const rssi = item.rssi;

        if (typeof rssi === 'number') {
          this.recordings.push({
            gatewayMac,
            rssi,
            timestamp: Date.now()
          });
        }
      });
    } catch (error) {
      // Ignore parse errors
    }
  }

  subscribe() {
    // Subscribe to all topics (mosquitto-client format)
    this.client.subscribe('#', (err) => {
      if (err) {
        console.error(`Error subscribing to topics:`, err);
      } else {
        console.log(`✓ Subscribed to all topics (#)`);
      }
    });
  }

  question(prompt) {
    return new Promise((resolve) => {
      this.rl.question(prompt, resolve);
    });
  }

  async recordMeasurement() {
    console.log('\n=== Gateway Calibration Recording ===\n');

    const gatewayMac = await this.question('Enter Gateway MAC address: ');
    if (!gatewayMac.trim()) {
      console.log('Gateway MAC is required. Skipping...\n');
      return false;
    }

    const distanceInput = await this.question('Enter distance from gateway (meters): ');
    const distance = parseFloat(distanceInput);
    if (isNaN(distance) || distance <= 0) {
      console.log('Invalid distance. Must be a positive number. Skipping...\n');
      return false;
    }

    console.log(`\nRecording RSSI for gateway ${gatewayMac} at ${distance}m distance...`);
    console.log('Recording for 1 minute. Please ensure device is at the specified distance.\n');

    // Reset recordings
    this.recordings = [];
    this.isRecording = true;
    this.currentGatewayMac = gatewayMac.trim().toUpperCase();
    this.currentDistance = distance;
    this.recordingStartTime = Date.now();

    // Show progress
    const progressInterval = setInterval(() => {
      const elapsed = Math.floor((Date.now() - this.recordingStartTime) / 1000);
      const remaining = 60 - elapsed;
      if (remaining > 0) {
        process.stdout.write(`\rRecording... ${elapsed}s / 60s (${this.recordings.length} samples)`);
      }
    }, 1000);

    // Wait for 1 minute
    await new Promise(resolve => setTimeout(resolve, RECORDING_DURATION));

    clearInterval(progressInterval);
    this.isRecording = false;

    // Calculate average
    if (this.recordings.length === 0) {
      console.log('\n\n⚠ No RSSI readings received during recording period.');
      console.log('Please check:');
      console.log('  1. MQTT broker is running');
      console.log('  2. Device is publishing RSSI data');
      console.log('  3. Gateway MAC address matches\n');
      return false;
    }

    const rssiValues = this.recordings.map(r => r.rssi);
    const avgRssi = rssiValues.reduce((a, b) => a + b, 0) / rssiValues.length;
    const minRssi = Math.min(...rssiValues);
    const maxRssi = Math.max(...rssiValues);

    console.log(`\n\n✓ Recording complete!`);
    console.log(`   Samples collected: ${this.recordings.length}`);
    console.log(`   Average RSSI: ${avgRssi.toFixed(2)} dBm`);
    console.log(`   Min RSSI: ${minRssi.toFixed(2)} dBm`);
    console.log(`   Max RSSI: ${maxRssi.toFixed(2)} dBm\n`);

    // Write to Excel
    await this.writeToExcel(this.currentGatewayMac, this.currentDistance, avgRssi);

    return true;
  }

  async writeToExcel(gatewayMac, distance, rssi) {
    const filePath = this.outputFile || path.join(__dirname, '..', 'gateway-calibration-data.xlsx');

    // Ensure directory exists
    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }

    let workbook;
    let data;
    const sheetName = 'Calibration Data';

    // Read existing file or create new
    if (fs.existsSync(filePath)) {
      workbook = XLSX.readFile(filePath);
      
      // Find existing sheet or use first sheet
      const existingSheetName = workbook.SheetNames.find(name => 
        name.toLowerCase().includes('calibration') || 
        name.toLowerCase().includes('data')
      ) || workbook.SheetNames[0];
      
      const worksheet = workbook.Sheets[existingSheetName];
      
      // Read all data including empty cells, preserving all rows
      data = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1, 
        defval: '',  // Default value for empty cells
        raw: false   // Convert values to strings/numbers
      });
      
      // Ensure we have at least headers
      if (data.length === 0) {
        data = [['Gateway MAC', 'Distance (m)', 'RSSI (dBm)', 'Notes', 'Timestamp']];
      }
      
      // If first row doesn't look like headers, add them
      if (data.length > 0 && !data[0].includes('Gateway MAC')) {
        data.unshift(['Gateway MAC', 'Distance (m)', 'RSSI (dBm)', 'Notes', 'Timestamp']);
      }
    } else {
      workbook = XLSX.utils.book_new();
      data = [['Gateway MAC', 'Distance (m)', 'RSSI (dBm)', 'Notes', 'Timestamp']];
    }

    // Append new row
    const newRow = [
      gatewayMac,
      distance,
      Math.round(rssi * 100) / 100, // Round to 2 decimals
      '',
      new Date().toISOString()
    ];
    data.push(newRow);

    // Create new worksheet from updated data
    const newWorksheet = XLSX.utils.aoa_to_sheet(data);
    
    // Set column widths
    newWorksheet['!cols'] = [
      { wch: 18 }, // Gateway MAC
      { wch: 12 }, // Distance
      { wch: 12 }, // RSSI
      { wch: 40 }, // Notes
      { wch: 25 }  // Timestamp
    ];

    // Update or add sheet to workbook
    if (workbook.SheetNames.includes(sheetName)) {
      workbook.Sheets[sheetName] = newWorksheet;
    } else {
      XLSX.utils.book_append_sheet(workbook, newWorksheet, sheetName);
    }

    // Write file
    XLSX.writeFile(workbook, filePath);
    console.log(`✓ Data saved to: ${filePath} (${data.length - 1} total entries)\n`);
  }

  async run() {
    try {
      console.log('=== Gateway Calibration Tool ===\n');
      console.log('This tool will:');
      console.log('  1. Connect to MQTT broker');
      console.log('  2. Record RSSI at specified distances');
      console.log('  3. Average readings over 1 minute');
      console.log('  4. Save to Excel file\n');

      const outputFileInput = await this.question(`Output file path (press Enter for default): `);
      if (outputFileInput.trim()) {
        this.outputFile = outputFileInput.trim();
      }

      await this.connect();
      this.subscribe();

      console.log('\nReady to record calibration data.\n');

      while (true) {
        const continueRecording = await this.question('Record another measurement? (y/n): ');
        if (continueRecording.toLowerCase() !== 'y') {
          break;
        }

        await this.recordMeasurement();
      }

      console.log('\n✓ Calibration session complete!');
      console.log(`Data saved to: ${this.outputFile || path.join(__dirname, '..', 'gateway-calibration-data.xlsx')}\n`);

    } catch (error) {
      console.error('\nError:', error.message);
      process.exit(1);
    } finally {
      if (this.client) {
        this.client.end();
      }
      this.rl.close();
    }
  }
}

// Run if executed directly
if (require.main === module) {
  const tool = new GatewayCalibrationTool();
  tool.run().catch(error => {
    console.error('Fatal error:', error);
    process.exit(1);
  });
}

module.exports = GatewayCalibrationTool;
