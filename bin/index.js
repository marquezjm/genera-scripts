#!/usr/bin/env node
import ExcelJS from 'exceljs';
import { Command } from 'commander';
import fs from 'fs';
import path from 'path';
import chalk from 'chalk';
import inquirer from 'inquirer';

const program = new Command();

program
  .version('1.0.1')
  .description('Generate SQL scripts from Excel file.')
  .argument('<inputFile>', 'Path to the input Excel file')
  .option('-k, --key-columns <keys>', 'Key columns configuration', 'NUMERO_INSTITUCION')
  .action(async (inputFile, options) => {
    try {
        await processFile(inputFile, options.keyColumns);
    } catch (error) {
        console.error(chalk.red('An unexpected error occurred:'), error);
        process.exit(1);
    }
  });

program.parse(process.argv);

function formatValue(val) {
  if (val === null || val === undefined) {
    return 'NULL';
  }
  
  // Handle objects (Formulas, Hyperlinks, RichText)
  if (typeof val === 'object' && !(val instanceof Date)) {
      if (val.result !== undefined) {
          // Formula result
          return formatValue(val.result);
      }
      if (val.richText) {
          // Rich Text
          return formatValue(val.richText.map(part => part.text).join(''));
      }
      if (val.text) {
          // Hyperlink
          return formatValue(val.text);
      }
  }

  // Handle dates
  if (val instanceof Date) {
      // Simple ISO format
      return `'${val.toISOString().replace('T', ' ').replace('Z', '')}'`;
  }

  let sVal = String(val).trim();
  
  if (sVal.toUpperCase() === 'NULL') {
    return 'NULL';
  }
  if (sVal.toUpperCase() === 'GETDATE()') {
    return 'GETDATE()';
  }
  
  // Escape single quotes
  return `'${sVal.replace(/'/g, "''")}'`;
}

function isCellStyled(cell) {
  // Check for Bold
  if (cell.font && cell.font.bold) {
    return true;
  }

  // Check for Fill Color (Background)
  if (cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid') {
    if (cell.fill.fgColor) {
         if (cell.fill.fgColor.argb && cell.fill.fgColor.argb !== '00000000') {
             return true;
         }
         if (cell.fill.fgColor.theme !== undefined) {
             return true;
         }
    }
  }
  return false;
}

async function processFile(inputFile, keyColumnsArg) {
  if (!fs.existsSync(inputFile)) {
    console.error(chalk.red(`Error: Input file not found at ${inputFile}`));
    process.exit(1);
  }

  // Parse key columns configuration
  let sheetKeysConfig = {};
  let globalDefaultKeys = ['NUMERO_INSTITUCION'];

  const rawConfig = keyColumnsArg;

  if (rawConfig.includes(':') || rawConfig.includes(';')) {
    const parts = rawConfig.split(';');
    for (let part of parts) {
      part = part.trim();
      if (!part) continue;

      if (part.includes(':')) {
        const [sName, cols] = part.split(':', 2);
        sheetKeysConfig[sName.trim()] = cols.split(',').map(c => c.trim());
      } else {
        globalDefaultKeys = part.split(',').map(c => c.trim());
      }
    }
  } else {
    globalDefaultKeys = rawConfig.split(',').map(c => c.trim());
  }

  console.log(chalk.blue(`Reading ${inputFile}...`));
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(inputFile);
  } catch (e) {
    console.error(chalk.red(`Error opening Excel file: ${e.message}`));
    process.exit(1);
  }

  const inputDir = path.dirname(inputFile);

  for (const sheet of workbook.worksheets) {
    const sheetName = sheet.name;
    const outputFilename = `${sheetName}.sql`;
    const outputPath = path.join(inputDir, outputFilename);

    console.log(chalk.cyan(`Processing sheet: ${sheetName}`));

    if (sheet.rowCount <= 0) {
        console.log(chalk.yellow("  -> Sheet is empty."));
        fs.writeFileSync(outputPath, `-- Sheet is empty\n`, 'utf-8');
        continue;
    }

    const headerRow = sheet.getRow(1);
    const headers = [];
    for(let i = 1; i <= headerRow.cellCount; i++) {
        const val = headerRow.getCell(i).value;
        headers.push(val ? String(val).trim() : '');
    }

    // Identify Action Column
    let actionColIdx = 0; 
    const possibleActionHeaders = ['ACCION', 'ACTION', 'TIPO', 'OPERACION', 'MOVIMIENTO'];
    
    let foundActionCol = false;
    for (let i = 0; i < headers.length; i++) {
        if (possibleActionHeaders.includes(headers[i].toUpperCase())) {
            actionColIdx = i;
            foundActionCol = true;
            break;
        }
    }

    console.log(`  -> Stats: ${sheet.rowCount} rows (approx).`);
    console.log(`  -> Action Column: '${headers[actionColIdx]}' (Index ${actionColIdx})`);

    let keyColumns = [...globalDefaultKeys];
    let deleteBeforeInsert = false;
    let insertKeyColumns = [...globalDefaultKeys];
    
    let hasUpdate = false;
    let hasInsert = false;
    
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        if (actionColIdx + 1 > row.cellCount) return;

        const cell = row.getCell(actionColIdx + 1);
        const val = cell.value ? String(cell.value).trim().toUpperCase() : '';
        if (val === 'UPDATE') {
            hasUpdate = true;
        }
        if (val === 'INSERT') {
            hasInsert = true;
        }
    });

    if (hasUpdate) {
        if (sheetKeysConfig[sheetName]) {
            keyColumns = sheetKeysConfig[sheetName];
            console.log(`  -> UPDATE detected. Using pre-configured keys: ${keyColumns.join(', ')}`);
        } else {
            const availableCols = headers.filter(h => h);
            
            console.log(chalk.yellow(`\n[!] UPDATE operations detected in '${sheetName}'`));
            
            const answers = await inquirer.prompt([
                {
                    type: 'checkbox',
                    name: 'selectedKeys',
                    message: 'Select the Key Columns for the WHERE clause:',
                    choices: availableCols,
                    default: globalDefaultKeys.filter(k => availableCols.includes(k)),
                    validate: (input) => {
                        if (input.length < 1) {
                            return 'You must choose at least one key column.';
                        }
                        return true;
                    }
                }
            ]);
            
            keyColumns = answers.selectedKeys;
            console.log(`    -> Selected keys: ${keyColumns.join(', ')}\n`);
        }
    }

    // Ask about DELETE before INSERT
    if (hasInsert) {
        const availableCols = headers.filter(h => h);
        
        console.log(chalk.yellow(`\n[!] INSERT operations detected in '${sheetName}'`));
        
        const deleteAnswer = await inquirer.prompt([
            {
                type: 'confirm',
                name: 'deleteBeforeInsert',
                message: 'Do you want to DELETE existing records before INSERT?',
                default: false
            }
        ]);
        
        deleteBeforeInsert = deleteAnswer.deleteBeforeInsert;
        
        if (deleteBeforeInsert) {
            console.log(chalk.cyan(`\n[!] Please select the Key Columns for the DELETE WHERE clause (based on INSERT data):`));
            
            const insertKeysAnswer = await inquirer.prompt([
                {
                    type: 'checkbox',
                    name: 'selectedKeys',
                    message: 'Select the Key Columns for DELETE WHERE:',
                    choices: availableCols,
                    default: globalDefaultKeys.filter(k => availableCols.includes(k)),
                    validate: (input) => {
                        if (input.length < 1) {
                            return 'You must choose at least one key column.';
                        }
                        return true;
                    }
                }
            ]);
            
            insertKeyColumns = insertKeysAnswer.selectedKeys;
            console.log(`    -> Selected keys for DELETE: ${insertKeyColumns.join(', ')}\n`);
        }
    }

    console.log(`Generating ${outputPath}...`);
    
    const colMap = {};
    headers.forEach((h, idx) => {
        if (h) colMap[h] = idx;
    });

    const hasKeys = keyColumns.every(k => Object.prototype.hasOwnProperty.call(colMap, k));
    const hasInsertKeys = deleteBeforeInsert ? insertKeyColumns.every(k => Object.prototype.hasOwnProperty.call(colMap, k)) : true;

    let deleteStatements = '';
    let insertStatements = '';
    let updateStatements = '';
    
    let generatedRows = 0;

    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        
        const actionCell = row.getCell(actionColIdx + 1);
        let actionVal = actionCell.value;
        if (typeof actionVal === 'object' && actionVal !== null && actionVal.result) actionVal = actionVal.result;
        
        const action = actionVal ? String(actionVal).trim().toUpperCase() : '';
        
        if (action !== 'INSERT' && action !== 'UPDATE') return;
        
        generatedRows++;

        if (action === 'INSERT') {
            // Generate DELETE before INSERT if requested
            if (deleteBeforeInsert) {
                if (!hasInsertKeys) {
                    const msg = `-- ERROR: Missing INSERT key columns ${insertKeyColumns.join(',')} in sheet ${sheetName}\n`;
                    deleteStatements += msg;
                    console.error(chalk.red(`    [!] Error: ${msg.trim()}`));
                    return;
                }
                
                const deleteWhereClauses = [];
                insertKeyColumns.forEach(key => {
                    const idx = colMap[key];
                    if (idx !== undefined) {
                        const cell = row.getCell(idx + 1);
                        const val = formatValue(cell.value);
                        deleteWhereClauses.push(`${key} = ${val}`);
                    }
                });
                
                if (deleteWhereClauses.length > 0) {
                    const whereStr = deleteWhereClauses.join(' AND ');
                    deleteStatements += `DELETE FROM ${sheetName} WHERE ${whereStr};\n`;
                }
            }
            
            // Generate INSERT
            const cols = [];
            const vals = [];
            
            headers.forEach((h, i) => {
                if (i === actionColIdx) return;
                if (!h) return;
                
                const cell = row.getCell(i + 1);
                cols.push(h);
                vals.push(formatValue(cell.value));
            });

            if (cols.length > 0) {
                const colStr = cols.join(', ');
                const valStr = vals.join(', ');
                insertStatements += `INSERT INTO ${sheetName} (${colStr}) VALUES (${valStr});\n`;
            }

        } else if (action === 'UPDATE') {
            if (!hasKeys) {
                const msg = `-- ERROR: Missing key columns ${keyColumns.join(',')} in sheet ${sheetName}\n`;
                updateStatements += msg;
                console.error(chalk.red(`    [!] Error: ${msg.trim()}`));
                return;
            }

            const setClauses = [];
            const whereClauses = [];
            
            keyColumns.forEach(key => {
                    const idx = colMap[key];
                    const cell = row.getCell(idx + 1);
                    const val = formatValue(cell.value);
                    whereClauses.push(`${key} = ${val}`);
            });

            headers.forEach((h, i) => {
                if (i === actionColIdx) return;
                if (!h) return;
                if (keyColumns.includes(h)) return;

                const cell = row.getCell(i + 1);
                if (isCellStyled(cell)) {
                    const val = formatValue(cell.value);
                    setClauses.push(`${h} = ${val}`);
                }
            });

            if (setClauses.length > 0) {
                    const setStr = setClauses.join(', ');
                    const whereStr = whereClauses.join(' AND ');
                    updateStatements += `UPDATE ${sheetName} SET ${setStr} WHERE ${whereStr};\n`;
            }
        }
    });

    // Combine all statements in order: DELETE -> INSERT -> UPDATE
    let sqlContent = `BEGIN TRAN TRAN_\n\nBEGIN TRY\n`;
    sqlContent += deleteStatements;
    sqlContent += insertStatements;
    sqlContent += updateStatements;

    if (generatedRows === 0) {
        console.log(chalk.yellow(`  -> Warning: No SQL statements were generated for '${sheetName}'. Check Action column values.`));
        sqlContent += "-- Warning: No valid INSERT/UPDATE rows found.\n";
    }

    sqlContent += `\nCOMMIT TRAN TRAN_;\n`;
    sqlContent += "PRINT 'PROCESO EJECUTADO CORRECTAMENTE';\n";
    sqlContent += "END TRY\n";
    sqlContent += "BEGIN CATCH\n";
    sqlContent += "    SELECT 'LINEA ERROR - ' + CAST(ERROR_LINE() AS VARCHAR(5)) + ': ' + ERROR_MESSAGE();\n";
    sqlContent += `    ROLLBACK TRAN TRAN_;\n`;
    sqlContent += "    PRINT 'OCURRIO UN ERROR EN EL PROCESO';\n";
    sqlContent += "END CATCH\n";

    fs.writeFileSync(outputPath, sqlContent, 'utf-8');
  }

  console.log(chalk.green("All scripts generated successfully."));
}
