#!/usr/bin/env node

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

class PriceListNormalizer {
  constructor() {
    // Price markup rules
    this.priceRules = [
      { min: 0, max: 10, add: 1.00 },
      { min: 10, max: 20, add: 1.50 },
      { min: 20, max: 30, add: 2.00 },
      { min: 30, max: 40, add: 2.50 },
      { min: 40, max: 50, add: 3.00 },
      { min: 50, max: Infinity, add: 3.50 }
    ];

    // Type normalization mapping
    this.typeMapping = {
      '100.Regular': 'Original Box',
      '101.Tester': 'Tester Box',
      '102.Set': 'Set',
      '103.Unbox': 'Unbox',
      '104.Mini': 'Mini',
      '105.Mini Set': 'Mini Set',
      '106.Body Mist': 'Body Mist',
      '107.Body Lotion': 'Body Lotion',
      '108.Deo': 'Deo',
      '109.Vial': 'Vial',
      '110.Candles': 'Candles',
      '111.Undefined': 'Undefined',
      '112.Body Wash': 'Body Wash',
      '113.Shower Gel': 'Shower Gel',
      '114.After Shave': 'After Shave',
      '115.Skincare and Cosmetics': 'Skincare and Cosmetics'
    };
  }

  applyPriceMarkup(price) {
    if (!price || isNaN(price)) return price;
    
    const numPrice = parseFloat(price);
    
    for (const rule of this.priceRules) {
      if (numPrice >= rule.min && numPrice < rule.max) {
        return parseFloat((numPrice + rule.add).toFixed(2));
      }
    }
    
    return parseFloat(numPrice.toFixed(2));
  }

  normalizeDescription(description) {
    if (!description) return '';
    
    let normalized = description.toString().trim();
    
    // Remove country codes like "- France -", "- Spain -"
    normalized = normalized.replace(/\s*-\s*[A-Z][a-z]+\s*-/g, '');
    
    // Remove product codes like "(116427)"
    normalized = normalized.replace(/\s*\(\d+\)/g, '');
    
    // Remove box quantity info like "- 57pcs ByBox"
    normalized = normalized.replace(/\s*-\s*\d+pcs\s+ByBox/gi, '');
    
    // Remove standalone "ByBox"
    normalized = normalized.replace(/\s*ByBox/gi, '');
    
    // Clean up multiple spaces
    normalized = normalized.replace(/\s+/g, ' ').trim();
    
    return normalized;
  }

  normalizeBrand(brand) {
    if (!brand) return '';
    return brand.toString().trim();
  }

  normalizeType(typeValue) {
    if (!typeValue) return 'Undefined';
    
    const typeStr = typeValue.toString().trim();
    return this.typeMapping[typeStr] || typeStr;
  }

  readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    // Convert to JSON array
    const data = XLSX.utils.sheet_to_json(worksheet, { 
      header: 1,
      defval: null,
      raw: false
    });
    
    return data;
  }

  process(inputFile, outputFile) {
    console.log(`Reading price list from: ${inputFile}`);
    const rawData = this.readExcel(inputFile);
    
    // Find the header row
    let headerRowIndex = -1;
    let headers = [];
    
    for (let i = 0; i < rawData.length; i++) {
      const row = rawData[i];
      const rowStr = row.join('|').toUpperCase();
      if (rowStr.includes('DESCRIPTION') && 
          rowStr.includes('PRICE') && 
          rowStr.includes('BRAND') && 
          rowStr.includes('TYPE')) {
        headerRowIndex = i;
        headers = row;
        break;
      }
    }
    
    if (headerRowIndex === -1) {
      throw new Error('Could not find header row with DESCRIPTION, PRICE, BRAND, and TYPE columns');
    }
    
    console.log(`Found headers at row ${headerRowIndex + 1}`);
    
    // Find column indices (trim spaces from headers)
    const descCol = headers.findIndex(h => h && h.toString().trim().toUpperCase().includes('DESCRIPTION'));
    const priceCol = headers.findIndex(h => h && h.toString().trim().toUpperCase() === 'PRICE');
    const brandCol = headers.findIndex(h => h && h.toString().trim().toUpperCase() === 'BRAND');
    const typeCol = headers.findIndex(h => h && h.toString().trim().toUpperCase() === 'TYPE');
    
    if (descCol === -1 || priceCol === -1 || brandCol === -1 || typeCol === -1) {
      throw new Error('Could not find required columns (DESCRIPTION, PRICE, BRAND, TYPE)');
    }
    
    console.log(`Column indices - Description: ${descCol}, Price: ${priceCol}, Brand: ${brandCol}, Type: ${typeCol}`);
    
    // Process data rows
    const processedData = [];
    processedData.push(['BRAND', 'DESCRIPTION', 'CAJA', 'PRICE']);
    
    let processedCount = 0;
    let skippedCount = 0;
    const priceStats = {};
    const typeStats = {};
    
    for (let i = headerRowIndex + 1; i < rawData.length; i++) {
      const row = rawData[i];
      
      const description = row[descCol];
      const price = row[priceCol];
      const brand = row[brandCol];
      const type = row[typeCol];
      
      // Skip rows without essential data
      if (!description || !price || !brand) {
        skippedCount++;
        continue;
      }
      
      const normalizedDesc = this.normalizeDescription(description);
      const normalizedBrand = this.normalizeBrand(brand);
      const normalizedType = this.normalizeType(type);
      const originalPrice = parseFloat(price);
      const newPrice = this.applyPriceMarkup(price);
      
      processedData.push([
        normalizedBrand,
        normalizedDesc,
        normalizedType,
        newPrice
      ]);
      
      // Collect stats
      for (const rule of this.priceRules) {
        if (originalPrice >= rule.min && originalPrice < rule.max) {
          const key = rule.max === Infinity ? `$${rule.min}+` : `$${rule.min}-$${rule.max}`;
          priceStats[key] = (priceStats[key] || 0) + 1;
          break;
        }
      }
      
      typeStats[normalizedType] = (typeStats[normalizedType] || 0) + 1;
      
      processedCount++;
    }
    
    console.log(`Processed ${processedCount} items, skipped ${skippedCount} rows`);
    
    // Write to new Excel file
    console.log(`Writing normalized price list to: ${outputFile}`);
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(processedData);
    
    // Set column widths
    newWorksheet['!cols'] = [
      { wch: 20 },  // BRAND
      { wch: 50 },  // DESCRIPTION
      { wch: 20 },  // CAJA
      { wch: 12 }   // PRICE
    ];
    
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Normalized Price List');
    XLSX.writeFile(newWorkbook, outputFile);
    
    console.log(`✓ Successfully created normalized price list`);
    console.log(`✓ Total products: ${processedCount}`);
    
    // Show summaries
    this.showPricingSummary(priceStats);
    this.showTypeSummary(typeStats);
    
    return processedData;
  }

  showPricingSummary(stats) {
    console.log('\nPricing Summary:');
    console.log('='.repeat(50));
    
    const order = ['$0-$10', '$10-$20', '$20-$30', '$30-$40', '$40-$50', '$50+'];
    const markups = [1.00, 1.50, 2.00, 2.50, 3.00, 3.50];
    
    order.forEach((range, idx) => {
      const count = stats[range] || 0;
      if (count > 0) {
        console.log(`${range.padEnd(12)} ${count.toString().padStart(5)} items (+$${markups[idx].toFixed(2)} each)`);
      }
    });
  }

  showTypeSummary(stats) {
    console.log('\nBox Type Summary:');
    console.log('='.repeat(50));
    
    // Sort by count descending
    const sorted = Object.entries(stats).sort((a, b) => b[1] - a[1]);
    
    sorted.forEach(([type, count]) => {
      console.log(`${type.padEnd(30)} ${count.toString().padStart(5)} items`);
    });
  }
}

// CLI usage
if (require.main === module) {
  const args = process.argv.slice(2);
  
  if (args.length < 1) {
    console.log('Usage: node normalize.js <input.xlsx> [output.xlsx]');
    console.log('\nExample: node normalize.js MTZpricelist.xlsx normalized-prices.xlsx');
    console.log('\nPrice markup rules:');
    console.log('  < $10:  +$1.00');
    console.log('  $10-20: +$1.50');
    console.log('  $20-30: +$2.00');
    console.log('  $30-40: +$2.50');
    console.log('  $40-50: +$3.00');
    console.log('  $50+:   +$3.50');
    console.log('\nBox type normalization:');
    console.log('  100.Regular → Original Box');
    console.log('  101.Tester  → Tester Box');
    console.log('  102.Set     → Set');
    console.log('  And more...');
    process.exit(1);
  }
  
  const inputFile = args[0];
  const outputFile = args[1] || 'normalized-prices.xlsx';
  
  if (!fs.existsSync(inputFile)) {
    console.error(`Error: File not found: ${inputFile}`);
    process.exit(1);
  }
  
  try {
    const normalizer = new PriceListNormalizer();
    normalizer.process(inputFile, outputFile);
  } catch (error) {
    console.error('Error:', error.message);
    console.error(error.stack);
    process.exit(1);
  }
}

module.exports = PriceListNormalizer;