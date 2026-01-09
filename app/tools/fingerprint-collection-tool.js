/**
 * Fingerprint Collection Tool
 * Listens to MQTT RSSI data, records at specific locations for 1 minute,
 * averages readings from all gateways, and writes to Excel file
 * Adapted for mosquitto-client message format
 */

const mqtt = require('mqtt');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const brokerUrl = process.env.MQTT_BROKER_URL || 'mqtt://localhost:1883';
const RECORDING_DURATION = 60 * 1000; // 1 minute in milliseconds

class FingerprintCollectionTool {
  constructor() {
    this.client = null;
    this.rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
    this.recordings = new Map(); // gatewayMac -> [rssi values]
    this.isRecording = false;
    this.recordingStartTime = null;
    this.currentLocationId = null;
    this.currentCoordinates = null;
    this.outputFile = null;
    this.gatewayMacs = new Set();
  }

  async connect() {
    return new Promise((resolve, reject) => {
      this.client = mqtt.connect(brokerUrl, {
        clientId: `fingerprint-tool-${Date.now()}`
      });

      this.client.on('connect', () => {
        console.log(`✓ Connected to MQTT broker: ${brokerUrl}\n`);
        this.   client.subscribe('#', (err) => {
          if (!err) {
              console.log('📡 Subscribed to all topics');
          } else {
              console.error('❌ Subscription error:', err);
          }
      });
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
      this.gatewayMacs.add(gatewayMac);

      // Collect RSSI values from all tags detected by this gateway
      payload.data.forEach(item => {
        const rssi = item.rssi;

        if (typeof rssi === 'number') {
          if (!this.recordings.has(gatewayMac)) {
            this.recordings.set(gatewayMac, []);
          }
          
          this.recordings.get(gatewayMac).push({
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
    console.log('\n=== Fingerprint Collection Recording ===\n');

    const locationId = await this.question('Enter Location ID (e.g., point-2-3): ');
    if (!locationId.trim()) {
      console.log('Location ID is required. Skipping...\n');
      return false;
    }

    const xInput = await this.question('Enter X coordinate (meters): ');
    const x = parseFloat(xInput);
    if (isNaN(x)) {
      console.log('Invalid X coordinate. Skipping...\n');
      return false;
    }

    const yInput = await this.question('Enter Y coordinate (meters): ');
    const y = parseFloat(yInput);
    if (isNaN(y)) {
      console.log('Invalid Y coordinate. Skipping...\n');
      return false;
    }

    const zInput = await this.question('Enter Z coordinate (meters, optional, press Enter for 0): ');
    const z = zInput.trim() ? parseFloat(zInput) : 0;
    if (isNaN(z)) {
      console.log('Invalid Z coordinate, using 0. Skipping...\n');
      return false;
    }

    console.log(`\nRecording RSSI from all gateways at location ${locationId} (${x}, ${y}, ${z})...`);
    console.log('Recording for 1 minute. Please ensure device is at the specified location.\n');

    // Reset recordings
    this.recordings.clear();
    this.gatewayMacs.clear();
    this.isRecording = true;
    this.currentLocationId = locationId.trim();
    this.currentCoordinates = { x, y, z };
    this.recordingStartTime = Date.now();

    // Show progress
    const progressInterval = setInterval(() => {
      const elapsed = Math.floor((Date.now() - this.recordingStartTime) / 1000);
      const remaining = 60 - elapsed;
      const totalSamples = Array.from(this.recordings.values())
        .reduce((sum, arr) => sum + arr.length, 0);
      if (remaining > 0) {
        process.stdout.write(`\rRecording... ${elapsed}s / 60s (${totalSamples} samples, ${this.gatewayMacs.size} gateways)`);
      }
    }, 1000);

    // Wait for 1 minute
    await new Promise(resolve => setTimeout(resolve, RECORDING_DURATION));

    clearInterval(progressInterval);
    this.isRecording = false;

    // Calculate averages per gateway
    if (this.recordings.size === 0) {
      console.log('\n\n⚠ No RSSI readings received during recording period.');
      console.log('Please check:');
      console.log('  1. MQTT broker is running');
      console.log('  2. Device is publishing RSSI data');
      console.log('  3. Gateways are active\n');
      return false;
    }

    const rssiReadings = {};
    const stats = {};

    this.recordings.forEach((values, gatewayMac) => {
      const rssiValues = values.map(v => v.rssi);
      const avgRssi = rssiValues.reduce((a, b) => a + b, 0) / rssiValues.length;
      const minRssi = Math.min(...rssiValues);
      const maxRssi = Math.max(...rssiValues);

      rssiReadings[gatewayMac] = Math.round(avgRssi * 100) / 100;
      stats[gatewayMac] = {
        samples: rssiValues.length,
        avg: avgRssi,
        min: minRssi,
        max: maxRssi
      };
    });

    console.log(`\n\n✓ Recording complete!`);
    console.log(`   Gateways detected: ${this.gatewayMacs.size}`);
    Object.entries(stats).forEach(([mac, stat]) => {
      console.log(`   ${mac}: ${stat.samples} samples, avg: ${stat.avg.toFixed(2)} dBm (${stat.min.toFixed(2)} to ${stat.max.toFixed(2)})`);
    });
    console.log();

    // Write to Excel
    await this.writeToExcel(this.currentLocationId, this.currentCoordinates, rssiReadings);

    return true;
  }

  async writeToExcel(locationId, coordinates, rssiReadings) {
    const filePath = this.outputFile || path.join(__dirname, '..', 'fingerprint-collection-data.xlsx');

    // Ensure directory exists
    const dir = path.dirname(filePath);
    if (!fs.existsSync(dir)) {
      fs.mkdirSync(dir, { recursive: true });
    }

    let workbook;
    let worksheet;
    let data;
    let headers;

    // Read existing file or create new
    if (fs.existsSync(filePath)) {
      workbook = XLSX.readFile(filePath);
      const existingSheetName = workbook.SheetNames.find(name => 
        name.toLowerCase().includes('fingerprint') || 
        name.toLowerCase().includes('data')
      ) || workbook.SheetNames[0] || 'Fingerprint Data';
      
      worksheet = workbook.Sheets[existingSheetName];
      
      // Read all data including empty cells, preserving all rows
      data = XLSX.utils.sheet_to_json(worksheet, { 
        header: 1,
        defval: '',  // Default value for empty cells
        raw: false   // Convert values to strings/numbers
      });
      
      headers = data[0] || [];
      
      // Ensure we have headers
      if (data.length === 0 || headers.length === 0) {
        headers = ['Location ID', 'X (m)', 'Y (m)', 'Z (m)'];
        data = [headers];
      }
    } else {
      workbook = XLSX.utils.book_new();
      headers = ['Location ID', 'X (m)', 'Y (m)', 'Z (m)'];
      data = [headers];
    }

    // Ensure all gateway columns exist in headers
    const gatewayMacs = Array.from(this.gatewayMacs).sort();
    gatewayMacs.forEach(mac => {
      if (!headers.includes(mac)) {
        headers.push(mac);
      }
    });

    // Build new row
    const newRow = [locationId, coordinates.x, coordinates.y, coordinates.z];
    gatewayMacs.forEach(mac => {
      newRow.push(rssiReadings[mac] || '');
    });

    // Update header row if needed
    if (data.length === 0 || data[0].length !== headers.length) {
      data[0] = headers;
    }

    // Append new row
    data.push(newRow);

    // Create new worksheet from updated data
    const newWorksheet = XLSX.utils.aoa_to_sheet(data);
    
    // Set column widths
    const colWidths = [
      { wch: 15 }, // Location ID
      { wch: 10 }, // X
      { wch: 10 }, // Y
      { wch: 10 }, // Z
      ...gatewayMacs.map(() => ({ wch: 18 })) // RSSI columns (MAC addresses)
    ];
    newWorksheet['!cols'] = colWidths;

    // Update or add sheet to workbook
    const sheetName = 'Fingerprint Data';
    if (workbook.SheetNames.includes(sheetName)) {
      workbook.Sheets[sheetName] = newWorksheet;
    } else {
      XLSX.utils.book_append_sheet(workbook, newWorksheet, sheetName);
    }

    // Write file
    XLSX.writeFile(workbook, filePath);
    console.log(`✓ Data saved to: ${filePath} (${data.length - 1} total locations)\n`);
  }

  async run() {
    try {
      console.log('=== Fingerprint Collection Tool ===\n');
      console.log('This tool will:');
      console.log('  1. Connect to MQTT broker');
      console.log('  2. Record RSSI from all gateways at specified locations');
      console.log('  3. Average readings over 1 minute');
      console.log('  4. Save to Excel file\n');

      const outputFileInput = await this.question(`Output file path (press Enter for default): `);
      if (outputFileInput.trim()) {
        this.outputFile = outputFileInput.trim();
      }

      await this.connect();
      this.subscribe();

      console.log('\nReady to record fingerprint data.\n');

      while (true) {
        const continueRecording = await this.question('Record another location? (y/n): ');
        if (continueRecording.toLowerCase() !== 'y') {
          break;
        }

        await this.recordMeasurement();
      }

      console.log('\n✓ Fingerprint collection session complete!');
      console.log(`Data saved to: ${this.outputFile || path.join(__dirname, '..', 'fingerprint-collection-data.xlsx')}\n`);

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
  const tool = new FingerprintCollectionTool();
  tool.run().catch(error => {
    console.error('Fatal error:', error);
    process.exit(1);
  });
}

module.exports = FingerprintCollectionTool;
