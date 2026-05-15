import * as XLSX from 'xlsx';

/**
 * Parses the messy Excel format (blocks per product) and converts
 * it to the clean flat-table standard.
 *
 * Original format pattern (repeats):
 *   Row A (info):    ProductName | _ | _ | Class(Wine) | Category(Red Wine) | _ | Density(98 g / 100 ml)
 *   Row B (headers): Size | Barcode | Code | Empty Wt. | Full Wt. | Packaging
 *   Row C (data 1):  750 ml / BOTTLE | 8423014951038 | - | 746 g | 1481 g | -
 *   Row D (data 2):  375 ml / BOTTLE | 5551234567890 | - | 300 g | 670 g  | -
 *   ...more data rows until next product info row...
 *
 * Target format:
 *   Codigo de barras | Clase | Categoria | Nombre de producto | Medida | Densidad | Peso Vacio | Peso lleno
 */
export const parseExcel = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', raw: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        const rows = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          defval: '',
          raw: true,
        });

        // Step 1: Identify all header rows (contain "Empty Wt.")
        const headerIndices = [];
        for (let i = 0; i < rows.length; i++) {
          if (isHeaderRow(rows[i])) {
            headerIndices.push(i);
          }
        }

        const cleanedData = [];

        // Step 2: For each header row, extract info row + ALL data rows below it
        for (let h = 0; h < headerIndices.length; h++) {
          const headerIdx = headerIndices[h];
          const headerRow = rows[headerIdx];

          // Info row is the one directly above the header
          const infoRow = headerIdx > 0 ? rows[headerIdx - 1] : [];

          // Build column map from this header
          const colMap = buildColumnMap(headerRow);

          // Extract product info (name, class, category, density)
          const productInfo = extractProductInfo(infoRow);

          // Data rows start right after header and go until next info row
          // The next info row would be at (next header index - 1)
          const nextHeaderIdx = h + 1 < headerIndices.length ? headerIndices[h + 1] : rows.length;
          // Data rows are from headerIdx+1 to nextHeaderIdx-2 (the row before the next info row)
          // But if there's no next header, go to end of file
          const dataEndIdx = h + 1 < headerIndices.length ? nextHeaderIdx - 1 : rows.length;

          for (let d = headerIdx + 1; d < dataEndIdx; d++) {
            const dataRow = rows[d];
            if (!dataRow || dataRow.length === 0) continue;

            // Skip if this row looks like another header or info row
            if (isHeaderRow(dataRow)) continue;
            if (isEmptyRow(dataRow)) continue;

            const record = extractDataRow(dataRow, colMap, productInfo);
            if (record) {
              cleanedData.push(record);
            }
          }
        }

        resolve(cleanedData);
      } catch (error) {
        console.error('Parse error:', error);
        reject(error);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

/**
 * Checks if a row is a header row by looking for "Empty Wt" in any cell.
 */
function isHeaderRow(row) {
  if (!row || row.length === 0) return false;
  for (let j = 0; j < row.length; j++) {
    const cell = String(row[j]).toLowerCase().trim();
    if (cell.includes('empty wt')) return true;
  }
  return false;
}

/**
 * Checks if a row is effectively empty (all cells blank or just dashes).
 */
function isEmptyRow(row) {
  if (!row) return true;
  return row.every((cell) => {
    const val = String(cell).trim();
    return val === '' || val === '-';
  });
}

/**
 * Builds a map of column names to indices from a header row.
 */
function buildColumnMap(headerRow) {
  const colMap = {};
  for (let j = 0; j < headerRow.length; j++) {
    const h = String(headerRow[j]).toLowerCase().trim();
    if (h.includes('size')) colMap.size = j;
    if (h.includes('barcode')) colMap.barcode = j;
    if (h.includes('code') && !h.includes('barcode')) colMap.code = j;
    if (h.includes('empty')) colMap.emptyWt = j;
    if (h.includes('full')) colMap.fullWt = j;
    if (h.includes('packaging') || h.includes('pack')) colMap.packaging = j;
  }
  return colMap;
}

/**
 * Extracts product-level info from the info row (name, class, category, density).
 */
function extractProductInfo(infoRow) {
  let productName = '';
  let clase = '';
  let categoria = '';
  let densidad = 0;

  const infoValues = (infoRow || []).map((c) => String(c).trim());

  // First non-empty cell is the product name
  for (let j = 0; j < infoValues.length; j++) {
    const val = infoValues[j];
    if (val && val.length > 1) {
      productName = val;
      break;
    }
  }

  for (let j = 0; j < infoValues.length; j++) {
    const val = infoValues[j];
    const lower = val.toLowerCase();

    // Skip the product name we already found
    if (val === productName && j === infoValues.indexOf(productName)) continue;

    // Detect density pattern: "98 g / 100 ml"
    const densityMatch = val.match(/(\d+(?:\.\d+)?)\s*g\s*\/\s*\d+\s*m[lL]/);
    if (densityMatch) {
      densidad = parseFloat(densityMatch[1]);
      continue;
    }

    // Detect class
    const classNames = [
      'wine', 'liquor', 'beer', 'champagne', 'brandy', 'bourbon',
      'tequila', 'vodka', 'rum', 'whiskey', 'whisky', 'gin', 'mezcal',
      'cognac', 'scotch', 'sake',
    ];
    if (classNames.includes(lower)) {
      clase = capitalizeFirst(lower);
      continue;
    }

    // Detect category keywords
    const categoryKeywords = [
      'red wine', 'white wine', 'rose wine', 'rosé wine', 'sparkling',
      'imported', 'domestic', 'blanco', 'reposado', 'añejo', 'cristalino',
      'extra añejo', 'joven',
    ];
    const isCategory = categoryKeywords.some((kw) => lower.includes(kw));
    if (isCategory) {
      categoria = val;
      continue;
    }

    // If we found class but no category yet, and this isn't a known skip word
    if (clase && !categoria && val && val !== productName) {
      const skipWords = ['unknown', 'default', 'imported', 'domestic'];
      if (!skipWords.includes(lower)) {
        categoria = val;
      }
    }
  }

  if (!clase) clase = 'Miscellaneous';
  if (!categoria) categoria = 'General';

  return { productName, clase, categoria, densidad };
}

/**
 * Extracts a single record from a data row using the column map and product info.
 */
function extractDataRow(dataRow, colMap, productInfo) {
  // Extract barcode — handle scientific notation
  let barcode = '';
  if (colMap.barcode !== undefined) {
    const rawBarcode = dataRow[colMap.barcode];
    if (typeof rawBarcode === 'number') {
      barcode = rawBarcode.toFixed(0);
    } else {
      const str = String(rawBarcode).trim();
      if (str && str !== '-') {
        if (str.includes('E') || str.includes('e')) {
          barcode = parseFloat(str).toFixed(0);
        } else {
          barcode = str;
        }
      }
    }
  }

  // Size / Medida — extract ml from "750 ml / BOTTLE"
  let medida = 0;
  if (colMap.size !== undefined) {
    const sizeVal = dataRow[colMap.size];
    const sizeStr = String(sizeVal).trim();
    const sizeMatch = sizeStr.match(/(\d+)\s*ml/i);
    if (sizeMatch) {
      medida = parseInt(sizeMatch[1], 10);
    }
  }

  // Weights
  const pesoVacio = parseWeight(dataRow[colMap.emptyWt]);
  const pesoLleno = parseWeight(dataRow[colMap.fullWt]);

  // If both weights are 0, skip this row (probably not a real data row)
  if (pesoVacio === 0 && pesoLleno === 0 && !barcode && medida === 0) {
    return null;
  }

  const densidad = productInfo.densidad;

  // Calculate medida from weights + density if not found in Size column
  if (medida === 0 && densidad > 0 && pesoLleno > 0 && pesoVacio > 0) {
    const pesoNeto = pesoLleno - pesoVacio;
    medida = Math.round(pesoNeto / (densidad / 100));
    const standards = [330, 355, 375, 500, 650, 700, 750, 1000, 1500, 1750, 3000];
    const closest = standards.reduce((prev, curr) =>
      Math.abs(curr - medida) < Math.abs(prev - medida) ? curr : prev
    );
    if (Math.abs(closest - medida) < 15) medida = closest;
  }

  // Calculate density if missing
  let densidadFinal = densidad;
  if (densidadFinal === 0 && medida > 0 && pesoLleno > 0 && pesoVacio > 0) {
    const pesoNeto = pesoLleno - pesoVacio;
    densidadFinal = Math.round((pesoNeto / medida) * 100 * 100) / 100;
  }

  return {
    'Codigo de barras': barcode,
    Clase: productInfo.clase,
    Categoria: productInfo.categoria,
    'Nombre de producto': productInfo.productName,
    Medida: medida || '',
    Densidad: densidadFinal || '',
    'Peso Vacio': pesoVacio || '',
    'Peso lleno': pesoLleno || '',
  };
}

/**
 * Parses a weight value like "746 g" or 746 into a number.
 */
function parseWeight(val) {
  if (val === undefined || val === null) return 0;
  if (typeof val === 'number') return val;
  const str = String(val).replace(/[^\d.]/g, '');
  return parseFloat(str) || 0;
}

function capitalizeFirst(str) {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

/**
 * Exports the cleaned data array to an Excel file and triggers download.
 */
export const exportToExcel = (data, filename = 'convertido.xlsx') => {
  const worksheet = XLSX.utils.json_to_sheet(data);

  // Auto-size columns
  const colWidths = Object.keys(data[0] || {}).map((key) => ({
    wch: Math.max(key.length, ...data.map((r) => String(r[key] || '').length)) + 2,
  }));
  worksheet['!cols'] = colWidths;

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  XLSX.writeFile(workbook, filename);
};
