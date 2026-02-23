const express = require('express');
const axios = require('axios');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json());

// ====================== HELPER FUNCTIONS ======================
function columnLetterToIndex(letter) {
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
        index = index * 26 + (letter.charCodeAt(i) - 64);
    }
    return index;
}

function addTable(worksheet, name, ref, styleName = 'TableStyleMedium2') {
    const match = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) throw new Error(`Invalid table reference: ${ref}`);

    const startColLetter = match[1];
    const startRow = parseInt(match[2], 10);
    const endColLetter = match[3];
    const endRow = parseInt(match[4], 10);

    // Check if there are any data rows (endRow > startRow)
    if (endRow <= startRow) {
        console.log(`Skipping table ${name} - no data rows`);
        return null;
    }

    const startCol = columnLetterToIndex(startColLetter);
    const endCol = columnLetterToIndex(endColLetter);

    const headerRow = worksheet.getRow(startRow);
    const columns = [];

    for (let col = startCol; col <= endCol; col++) {
        const cell = headerRow.getCell(col);
        const columnName = cell.value ? cell.value.toString() : `Column${col}`;
        columns.push({ name: columnName, filterButton: true });
    }

    // Verify that we have at least one data row by checking if any cell in the next row has a value
    let hasData = false;
    const dataRow = worksheet.getRow(startRow + 1);
    for (let col = startCol; col <= endCol; col++) {
        if (dataRow.getCell(col).value) {
            hasData = true;
            break;
        }
    }

    if (!hasData) {
        console.log(`Skipping table ${name} - no data in rows`);
        return null;
    }

    try {
        const table = worksheet.addTable({
            name: name.replace(/[^a-zA-Z0-9_]/g, '_'),
            ref: ref,
            columns: columns,
            style: { theme: styleName, showRowStripes: true },
        });
        return table;
    } catch (error) {
        console.log(`Error adding table ${name}:`, error.message);
        return null;
    }
}

async function autoFitColumns(worksheet) {
    worksheet.columns.forEach(column => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, cell => {
            const cellValue = cell.value ? cell.value.toString() : '';
            maxLength = Math.max(maxLength, cellValue.length);
        });
        column.width = Math.min(50, maxLength + 2);
    });
}

function ensureDir(dirPath) {
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
    }
}

function cleanServerUrl(server) {
    return server.replace(/^https?:\/\//, '').replace(/\/+$/, '').trim();
}

// ====================== KOBO API FETCHERS ======================
async function fetchKoboForm(server, token, uid) {
    const cleanServer = cleanServerUrl(server);
    const url = `https://${cleanServer}/api/v2/assets/${uid}/`;
    console.log('Fetching form from:', url);

    try {
        const response = await axios.get(url, {
            headers: { Authorization: `Token ${token}` }
        });
        return response.data;
    } catch (error) {
        console.error('Error fetching form:', error.message);
        if (error.response) {
            console.error('Status:', error.response.status);
            console.error('Data:', error.response.data);
        }
        throw new Error(`Failed to fetch form: ${error.message}`);
    }
}

async function fetchKoboData(server, token, uid) {
    let submissions = [];
    const cleanServer = cleanServerUrl(server);
    let nextUrl = `https://${cleanServer}/api/v2/assets/${uid}/data/?format=json`;

    while (nextUrl) {
        console.log('Fetching data from:', nextUrl);
        try {
            const response = await axios.get(nextUrl, {
                headers: { Authorization: `Token ${token}` }
            });

            submissions = submissions.concat(response.data.results || []);
            nextUrl = response.data.next;

            console.log(`Fetched ${response.data.results?.length || 0} submissions. Total: ${submissions.length}`);
        } catch (error) {
            console.error('Error fetching data:', error.message);
            if (error.response) {
                console.error('Status:', error.response.status);
                console.error('Data:', error.response.data);
            }
            throw new Error(`Failed to fetch data: ${error.message}`);
        }
    }

    return submissions;
}

// ====================== PROCESSING FUNCTIONS ======================
async function processAttendancePerStudent(formJson, submissions) {
    console.log('\n--- PER-STUDENT ATTENDANCE ---');
    const BASE_DIR = 'الحضور';

    const TARGET_STAGES = [
        'الطفولة', 'الاعدادي', 'الثانوي', 'الجامعيين',
        'العاملين', 'السيدات', 'العائلات'
    ];

    const STAGE_DISPLAY = {
        'الطفولة': 'الطفولة',
        'الاعدادي': 'الإعدادي',
        'الثانوي': 'الثانوي',
        'الجامعيين': 'الجامعيين',
        'العاملين': 'العاملين',
        'السيدات': 'السيدات',
        'العائلات': 'العائلات'
    };

    const CLASS_DISPLAY = {
        'ملائكة': 'الملائكة',
        'اول': 'الأول',
        'ثاني': 'الثاني',
        'ثالث': 'الثالث',
        'رابع': 'الرابع',
        'خامس': 'الخامس',
        'سادس': 'السادس',
        'سابع': 'السابع',
        'ثامن': 'الثامن',
        'تاسع': 'التاسع',
        'عاشر': 'العاشر',
        'حادي عشر': 'الحادي عشر',
        'ثاني عشر': 'الثاني عشر'
    };

    // Build roster from form
    function buildRosterFromForm(form) {
        const survey = form.content.survey;
        const choices = form.content.choices;

        // Stage choices
        const stageCodeToDisplay = {};
        const stageDisplayToCode = {};
        const classCodeToDisplay = {};
        const classNameToCode = {};

        choices.forEach(c => {
            if (c.list_name === 'di1ez97') {
                stageCodeToDisplay[c.name] = c['label::Arabic (ar)'] || c.label;
                stageDisplayToCode[c['label::Arabic (ar)'] || c.label] = c.name;
            }
            if (['uc4lx56', 'sm64v07', 'ga6gx66'].includes(c.list_name)) {
                classCodeToDisplay[c.name] = c['label::Arabic (ar)'] || c.label;
                classNameToCode[c['label::Arabic (ar)'] || c.label] = c.name;
            }
        });

        const roster = {};
        let currentGroup = null;

        survey.forEach(row => {
            const type = row.type;
            if (type === 'begin_group') {
                const relevant = row.relevant;
                if (!relevant) return;

                const match = relevant.match(/\$\{([^}]+)\}\s*=\s*'([^']+)'/);
                if (match) {
                    const varName = match[1];
                    const val = match[2];

                    let stageCode, classCode;
                    if (varName === 'select_one_ew0hj33') {
                        stageCode = val;
                        classCode = val;
                    } else if (varName === 'select_one_ee9xi04') {
                        stageCode = '_1'; // Childhood
                        classCode = val;
                    } else if (varName === 'select_one_ks9wr72') {
                        stageCode = '_2'; // Intermediate
                        classCode = val;
                    } else if (varName === 'select_one_bl1bm11') {
                        stageCode = '_3'; // Secondary
                        classCode = val;
                    } else {
                        return;
                    }

                    currentGroup = { stageCode, classCode };
                    if (!roster[stageCode]) roster[stageCode] = {};
                    if (!roster[stageCode][classCode]) roster[stageCode][classCode] = new Set();
                }
            } else if (type === 'acknowledge' && currentGroup) {
                const student = row['label::Arabic (ar)'] || row.label;
                if (student) {
                    let studentName = '';
                    if (Array.isArray(student)) {
                        studentName = student[0] ? student[0].toString().trim() : '';
                    } else {
                        studentName = student.toString().trim();
                    }
                    if (studentName) {
                        roster[currentGroup.stageCode][currentGroup.classCode].add(studentName);
                    }
                }
            } else if (type === 'end_group') {
                currentGroup = null;
            }
        });

        return { stageCodeToDisplay, stageDisplayToCode, classCodeToDisplay, classNameToCode, roster };
    }

    console.log('Building roster from form JSON...');
    const formData = buildRosterFromForm(formJson);
    const { stageDisplayToCode, classNameToCode, roster } = formData;
    console.log(`Roster built for ${Object.keys(roster).length} stages.`);

    // Create directories
    ensureDir(BASE_DIR);
    TARGET_STAGES.forEach(stage => {
        const folder = STAGE_DISPLAY[stage] || stage;
        ensureDir(path.join(BASE_DIR, folder));
    });

    // Transform submissions into sessions
    const sessions = [];

    submissions.forEach(sub => {
        const dateStr = sub['date_ub3xq22'];
        if (!dateStr) return;
        // Keep the original string as the date key (YYYY-MM-DD)
        const dateKey = dateStr; // e.g., "2025-12-05"
        const dateObj = new Date(dateStr); // for sorting only

        const stageCode = sub['select_one_ew0hj33'];
        if (!stageCode) return;

        let stageDisplay = '';
        for (const [code, display] of Object.entries(formData.stageCodeToDisplay)) {
            if (code === stageCode) {
                stageDisplay = display;
                break;
            }
        }

        let classDisplay = stageDisplay;
        let classCode = null;

        if (stageCode === '_1' && sub['select_one_ee9xi04']) {
            classCode = sub['select_one_ee9xi04'];
            classDisplay = formData.classCodeToDisplay[classCode] || classCode;
        } else if (stageCode === '_2' && sub['select_one_ks9wr72']) {
            classCode = sub['select_one_ks9wr72'];
            classDisplay = formData.classCodeToDisplay[classCode] || classCode;
        } else if (stageCode === '_3' && sub['select_one_bl1bm11']) {
            classCode = sub['select_one_bl1bm11'];
            classDisplay = formData.classCodeToDisplay[classCode] || classCode;
        }

        const teacher = sub['text_fg36m00'] || '';
        const title = sub['text_pu3fc60'] || '';

        const presentStudents = new Set();

        // Collect attendance from all acknowledge fields with "OK"
        Object.keys(sub).forEach(key => {
            if (sub[key] === 'OK') {
                const parts = key.split('/');
                const fieldName = parts[parts.length - 1];
                const surveyItem = formJson.content.survey.find(item => item.name === fieldName);
                if (surveyItem) {
                    const studentName = surveyItem['label::Arabic (ar)'] || surveyItem.label;
                    if (studentName) {
                        let nameToAdd = '';
                        if (Array.isArray(studentName)) {
                            nameToAdd = studentName[0] ? studentName[0].toString().trim() : '';
                        } else {
                            nameToAdd = studentName.toString().trim();
                        }
                        if (nameToAdd) {
                            presentStudents.add(nameToAdd);
                        }
                    }
                }
            }
        });

        // New students from repeat group
        if (sub['group_ei9ze49'] && Array.isArray(sub['group_ei9ze49'])) {
            sub['group_ei9ze49'].forEach(item => {
                const newStudent = item['group_ei9ze49/text_bp9cv23'];
                if (newStudent) {
                    newStudent.split('\n').forEach(name => {
                        const trimmed = name.trim();
                        if (trimmed) presentStudents.add(trimmed);
                    });
                }
            });
        }

        sessions.push({
            dateKey,           // string YYYY-MM-DD
            dateObj,           // for sorting
            stage: stageDisplay,
            class: classDisplay,
            teacher,
            title,
            students: Array.from(presentStudents)
        });
    });

    console.log(`Extracted ${sessions.length} session records.`);

    // Group by stage -> class -> dateKey -> Set of students
    const stageData = {};
    sessions.forEach(sess => {
        const { stage, class: cls, dateKey, students } = sess;
        if (!stageData[stage]) stageData[stage] = {};
        if (!stageData[stage][cls]) stageData[stage][cls] = {};

        if (!stageData[stage][cls][dateKey]) {
            stageData[stage][cls][dateKey] = new Set();
        }
        students.forEach(s => stageData[stage][cls][dateKey].add(s));
    });

    // Helper: write one class-month file
    async function writeClassMonthFile(classDisplay, dates, allStudents, attendanceByDate, year, month, outputPath) {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet(classDisplay.slice(0, 31));

        let currentRow = 1;
        ws.getCell(`A${currentRow}`).value = `الصف : ${classDisplay}`;
        ws.getCell(`A${currentRow}`).font = { bold: true, size: 12 };
        currentRow++;

        const monthNames = ['يناير', 'فبراير', 'مارس', 'إبريل', 'مايو', 'يونيو',
            'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
        const headerText = `${monthNames[month - 1]} ${year}`;
        const colOffset = 3;
        const monthRow = currentRow;
        const dayRow = currentRow + 1;
        const studentStartRow = currentRow + 2;

        ws.mergeCells(monthRow, colOffset, monthRow, colOffset + dates.length - 1);
        ws.getCell(monthRow, colOffset).value = headerText;
        ws.getCell(monthRow, colOffset).font = { bold: true };
        ws.getCell(monthRow, colOffset).alignment = { horizontal: 'center' };

        for (let i = 0; i < dates.length; i++) {
            const day = dates[i].split('-')[2]; // extract day from YYYY-MM-DD
            ws.getCell(dayRow, colOffset + i).value = day;
            ws.getCell(dayRow, colOffset + i).font = { bold: true };
            ws.getCell(dayRow, colOffset + i).alignment = { horizontal: 'center' };
        }

        for (let i = 0; i < allStudents.length; i++) {
            const student = allStudents[i];
            const row = studentStartRow + i;
            ws.getCell(row, 1).value = i + 1;
            ws.getCell(row, 2).value = student;

            for (let j = 0; j < dates.length; j++) {
                const dateKey = dates[j]; // already string
                const presentSet = attendanceByDate[dateKey] || new Set();
                const status = presentSet.has(student) ? 'م' : 'غ';
                const cell = ws.getCell(row, colOffset + j);
                cell.value = status;
                cell.alignment = { horizontal: 'center' };

                if (status === 'م') {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } };
                    cell.font = { color: { argb: '006100' } };
                } else {
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };
                    cell.font = { color: { argb: '9C0006' } };
                }
            }
        }

        for (let r = monthRow; r <= studentStartRow + allStudents.length - 1; r++) {
            for (let c = 1; c <= colOffset + dates.length - 1; c++) {
                const cell = ws.getCell(r, c);
                cell.alignment = { horizontal: 'center', vertical: 'center' };
            }
        }

        await autoFitColumns(ws);

        const tableRef = `A${dayRow}:${String.fromCharCode(64 + colOffset + dates.length - 1)}${studentStartRow + allStudents.length - 1}`;
        addTable(ws, `Table_${classDisplay}_${year}_${month}`, tableRef);

        await wb.xlsx.writeFile(outputPath);
    }

    // Generate files
    for (const stage in stageData) {
        const stageFolder = STAGE_DISPLAY[stage] || stage;
        const stagePath = path.join(BASE_DIR, stageFolder);
        const stageCode = stageDisplayToCode[stage];

        for (const cls in stageData[stage]) {
            const dateAttendance = stageData[stage][cls];
            const displayClass = CLASS_DISPLAY[cls] || cls;

            let classCode = null;
            if (stageCode) {
                classCode = classNameToCode[cls];
                if (!classCode && cls === stage) classCode = stageCode;
            }

            const rosterStudents = (stageCode && classCode && roster[stageCode] && roster[stageCode][classCode])
                ? roster[stageCode][classCode]
                : new Set();

            const attendedStudents = new Set();
            for (const stuSet of Object.values(dateAttendance)) {
                stuSet.forEach(s => attendedStudents.add(s));
            }

            const allStudentsSet = new Set([...attendedStudents, ...rosterStudents]);
            if (allStudentsSet.size === 0) continue;

            const allStudents = Array.from(allStudentsSet).sort();

            const allDates = Object.keys(dateAttendance).sort();
            if (allDates.length === 0) continue;

            // Group by month using date string
            const monthsMap = {};
            for (const d of allDates) {
                const [year, month] = d.split('-').map(Number);
                const key = `${year}-${month}`;
                if (!monthsMap[key]) monthsMap[key] = [];
                monthsMap[key].push(d);
            }

            // Monthly files
            for (const [key, monthDates] of Object.entries(monthsMap)) {
                const [year, month] = key.split('-').map(Number);
                const monthAttendance = {};
                for (const d of monthDates) {
                    monthAttendance[d] = dateAttendance[d] || new Set();
                }

                const yearFolder = year.toString();
                const yearPath = path.join(stagePath, yearFolder);
                ensureDir(yearPath);
                const classPath = path.join(yearPath, displayClass);
                ensureDir(classPath);

                const filename = `حضور_${displayClass}_${year}-${month.toString().padStart(2, '0')}.xlsx`;
                const filepath = path.join(classPath, filename);
                await writeClassMonthFile(displayClass, monthDates, allStudents, monthAttendance, year, month, filepath);
                console.log(`Created: ${filepath}`);
            }

            // Full class file (all months together)
            const fullWb = new ExcelJS.Workbook();
            const fullWs = fullWb.addWorksheet(displayClass.slice(0, 31));
            let currentRow = 1;

            fullWs.getCell(`A${currentRow}`).value = `الصف : ${displayClass}`;
            fullWs.getCell(`A${currentRow}`).font = { bold: true, size: 12 };
            currentRow++;

            const colOffset = 3;
            const monthRow = currentRow;
            const dayRow = currentRow + 1;
            const studentStartRow = currentRow + 2;

            // Month headers
            let col = colOffset;
            for (const [key, monthDates] of Object.entries(monthsMap)) {
                const [year, month] = key.split('-').map(Number);
                const monthNames = ['يناير', 'فبراير', 'مارس', 'إبريل', 'مايو', 'يونيو',
                    'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
                const headerText = `${monthNames[month - 1]} ${year}`;
                fullWs.mergeCells(monthRow, col, monthRow, col + monthDates.length - 1);
                fullWs.getCell(monthRow, col).value = headerText;
                fullWs.getCell(monthRow, col).font = { bold: true };
                fullWs.getCell(monthRow, col).alignment = { horizontal: 'center' };
                col += monthDates.length;
            }

            // Day numbers
            col = colOffset;
            for (const monthDates of Object.values(monthsMap)) {
                for (const d of monthDates) {
                    const day = d.split('-')[2];
                    fullWs.getCell(dayRow, col).value = day;
                    fullWs.getCell(dayRow, col).font = { bold: true };
                    fullWs.getCell(dayRow, col).alignment = { horizontal: 'center' };
                    col++;
                }
            }

            // Student rows
            for (let i = 0; i < allStudents.length; i++) {
                const student = allStudents[i];
                const row = studentStartRow + i;
                fullWs.getCell(row, 1).value = i + 1;
                fullWs.getCell(row, 2).value = student;

                col = colOffset;
                for (const d of allDates) {
                    const presentSet = dateAttendance[d] || new Set();
                    const status = presentSet.has(student) ? 'م' : 'غ';
                    const cell = fullWs.getCell(row, col);
                    cell.value = status;
                    cell.alignment = { horizontal: 'center' };

                    if (status === 'م') {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'C6EFCE' } };
                        cell.font = { color: { argb: '006100' } };
                    } else {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC7CE' } };
                        cell.font = { color: { argb: '9C0006' } };
                    }
                    col++;
                }
            }

            for (let r = monthRow; r <= studentStartRow + allStudents.length - 1; r++) {
                for (let c = 1; c <= colOffset + allDates.length - 1; c++) {
                    const cell = fullWs.getCell(r, c);
                    cell.alignment = { horizontal: 'center', vertical: 'center' };
                }
            }

            await autoFitColumns(fullWs);
            const tableRef = `A${dayRow}:${String.fromCharCode(64 + colOffset + allDates.length - 1)}${studentStartRow + allStudents.length - 1}`;
            addTable(fullWs, `Table_${displayClass}_full`, tableRef);

            const fullFilename = `حضور_${displayClass}_الكامل.xlsx`;
            const fullFilepath = path.join(stagePath, fullFilename);
            await fullWb.xlsx.writeFile(fullFilepath);
            console.log(`Created full class file: ${fullFilepath}`);
        }
    }

    console.log('Per-student attendance processing finished.');
}

async function processAttendanceSessionCount(formJson, submissions) {
    console.log('\n--- SESSION-COUNT ATTENDANCE ---');
    const BASE_DIR = 'تفقد المواضيع';

    const TARGET_STAGES = [
        'الطفولة', 'الإعدادي', 'الثانوي', 'الجامعيين',
        'العاملين', 'العائلات', 'السيدات'
    ];

    const STAGE_MAP = {
        'الطفولة': 'الطفولة',
        'الاعدادي': 'الإعدادي',
        'الثانوي': 'الثانوي',
        'الجامعيين': 'الجامعيين',
        'العاملين': 'العاملين',
        'السيدات': 'السيدات',
        'العائلات': 'العائلات'
    };

    // Build student->class mapping
    function buildStudentToClass(form) {
        const survey = form.content.survey;
        const choices = form.content.choices;

        const classCodeToDisplay = {};
        choices.forEach(c => {
            if (['uc4lx56', 'sm64v07', 'ga6gx66'].includes(c.list_name)) {
                classCodeToDisplay[c.name] = c['label::Arabic (ar)'] || c.label;
            }
        });

        const studentToClass = {};
        let currentGroup = null;

        survey.forEach(row => {
            const type = row.type;
            if (type === 'begin_group') {
                const relevant = row.relevant;
                if (!relevant) return;

                const match = relevant.match(/\$\{([^}]+)\}\s*=\s*'([^']+)'/);
                if (match) {
                    const varName = match[1];
                    const val = match[2];

                    if (['select_one_ee9xi04', 'select_one_ks9wr72', 'select_one_bl1bm11'].includes(varName)) {
                        currentGroup = val; // class code
                    } else {
                        currentGroup = null;
                    }
                }
            } else if (type === 'acknowledge' && currentGroup) {
                const student = row['label::Arabic (ar)'] || row.label;
                if (student) {
                    let studentName = '';
                    if (Array.isArray(student)) {
                        studentName = student[0] ? student[0].toString().trim() : '';
                    } else {
                        studentName = student.toString().trim();
                    }
                    if (studentName) {
                        const classDisplay = classCodeToDisplay[currentGroup] || currentGroup;
                        studentToClass[studentName] = classDisplay;
                    }
                }
            } else if (type === 'end_group') {
                currentGroup = null;
            }
        });

        return studentToClass;
    }

    console.log('Building student->class mapping...');
    const studentToClass = buildStudentToClass(formJson);
    console.log(`Mapped ${Object.keys(studentToClass).length} students to classes.`);

    // Create directories
    ensureDir(BASE_DIR);
    TARGET_STAGES.forEach(stage => {
        ensureDir(path.join(BASE_DIR, stage));
    });

    // Transform submissions into records
    const records = [];

    submissions.forEach(sub => {
        const dateStr = sub['date_ub3xq22'];
        if (!dateStr) return;
        const dateKey = dateStr; // string YYYY-MM-DD
        const dateObj = new Date(dateStr); // for sorting

        const stageCode = sub['select_one_ew0hj33'];
        if (!stageCode) return;

        let stageDisplay = '';
        const choices = formJson.content.choices;
        const stageChoice = choices.find(c => c.list_name === 'di1ez97' && c.name === stageCode);
        if (stageChoice) {
            stageDisplay = stageChoice['label::Arabic (ar)'] || stageChoice.label;
        }

        const stageOut = STAGE_MAP[stageDisplay] || stageDisplay;

        const teacher = sub['text_fg36m00'] || '';
        const title = sub['text_pu3fc60'] || '';

        const presentStudents = new Set();

        Object.keys(sub).forEach(key => {
            if (sub[key] === 'OK') {
                const parts = key.split('/');
                const fieldName = parts[parts.length - 1];
                const surveyItem = formJson.content.survey.find(item => item.name === fieldName);
                if (surveyItem) {
                    const studentName = surveyItem['label::Arabic (ar)'] || surveyItem.label;
                    if (studentName) {
                        let nameToAdd = '';
                        if (Array.isArray(studentName)) {
                            nameToAdd = studentName[0] ? studentName[0].toString().trim() : '';
                        } else {
                            nameToAdd = studentName.toString().trim();
                        }
                        if (nameToAdd) {
                            presentStudents.add(nameToAdd);
                        }
                    }
                }
            }
        });

        // New students from repeat group
        if (sub['group_ei9ze49'] && Array.isArray(sub['group_ei9ze49'])) {
            sub['group_ei9ze49'].forEach(item => {
                const newStudent = item['group_ei9ze49/text_bp9cv23'];
                if (newStudent) {
                    newStudent.split('\n').forEach(name => {
                        const trimmed = name.trim();
                        if (trimmed) presentStudents.add(trimmed);
                    });
                }
            });
        }

        // Determine class(es) for this session
        let classVals = [];

        if (stageCode === '_1' && sub['select_one_ee9xi04']) {
            const classCode = sub['select_one_ee9xi04'];
            const classChoice = choices.find(c => c.list_name === 'uc4lx56' && c.name === classCode);
            if (classChoice) {
                classVals.push(classChoice['label::Arabic (ar)'] || classChoice.label);
            }
        } else if (stageCode === '_2' && sub['select_one_ks9wr72']) {
            const classCode = sub['select_one_ks9wr72'];
            const classChoice = choices.find(c => c.list_name === 'sm64v07' && c.name === classCode);
            if (classChoice) {
                classVals.push(classChoice['label::Arabic (ar)'] || classChoice.label);
            }
        } else if (stageCode === '_3' && sub['select_one_bl1bm11']) {
            const classCode = sub['select_one_bl1bm11'];
            const classChoice = choices.find(c => c.list_name === 'ga6gx66' && c.name === classCode);
            if (classChoice) {
                classVals.push(classChoice['label::Arabic (ar)'] || classChoice.label);
            }
        }

        if (classVals.length === 0) {
            // Infer class from attending students
            const classCounts = {};
            presentStudents.forEach(student => {
                const cls = studentToClass[student] || stageOut;
                classCounts[cls] = (classCounts[cls] || 0) + 1;
            });

            for (const [cls, cnt] of Object.entries(classCounts)) {
                records.push({
                    stage: stageOut,
                    dateKey,
                    dateObj,
                    class: cls,
                    count: cnt,
                    teacher,
                    title
                });
            }
        } else {
            // Use source classes, each gets full count
            classVals.forEach(cls => {
                records.push({
                    stage: stageOut,
                    dateKey,
                    dateObj,
                    class: cls,
                    count: presentStudents.size,
                    teacher,
                    title
                });
            });
        }
    });

    console.log(`Extracted ${records.length} session records.`);

    // Helper: auto-fit for session-count sheets
    async function autoFitSimple(ws) {
        ws.columns.forEach(col => {
            let maxLen = 0;
            col.eachCell({ includeEmpty: true }, cell => {
                const val = cell.value ? cell.value.toString() : '';
                maxLen = Math.max(maxLen, val.length);
            });
            col.width = Math.min(40, maxLen + 2);
        });
    }

    // Group records by stage, class, year-month using dateKey
    const stageClassMonth = {};
    for (const rec of records) {
        const { stage, class: cls, dateKey } = rec;
        const [year, month] = dateKey.split('-').map(Number);
        const key = `${year}-${month}`;

        if (!stageClassMonth[stage]) stageClassMonth[stage] = {};
        if (!stageClassMonth[stage][cls]) stageClassMonth[stage][cls] = {};
        if (!stageClassMonth[stage][cls][key]) stageClassMonth[stage][cls][key] = [];
        stageClassMonth[stage][cls][key].push(rec);
    }

    // Write monthly files
    for (const stage in stageClassMonth) {
        const stagePath = path.join(BASE_DIR, stage);
        for (const cls in stageClassMonth[stage]) {
            for (const [ym, recs] of Object.entries(stageClassMonth[stage][cls])) {
                const [year, month] = ym.split('-').map(Number);
                const sortedRecs = recs.sort((a, b) => a.dateObj - b.dateObj);
                const data = sortedRecs.map(r => ({
                    التاريخ: r.dateKey, // use original string
                    الصف: r.class,
                    عدد_الحضور: r.count,
                    حضور_المرشد: r.teacher,
                    عنوان_الموضوع: r.title
                }));

                const yearFolder = year.toString();
                const yearPath = path.join(stagePath, yearFolder);
                ensureDir(yearPath);
                const classPath = path.join(yearPath, cls);
                ensureDir(classPath);

                const wb = new ExcelJS.Workbook();
                const ws = wb.addWorksheet(cls.slice(0, 31));
                ws.columns = [
                    { header: 'التاريخ', key: 'التاريخ', width: 15 },
                    { header: 'الصف', key: 'الصف', width: 15 },
                    { header: 'عدد الحضور', key: 'عدد_الحضور', width: 15 },
                    { header: 'حضور المرشد', key: 'حضور_المرشد', width: 20 },
                    { header: 'عنوان الموضوع', key: 'عنوان_الموضوع', width: 30 }
                ];
                ws.addRows(data);

                ws.getRow(1).font = { bold: true };
                await autoFitSimple(ws);

                const tableRef = `A1:E${data.length + 1}`;
                addTable(ws, `Table_${cls}_${year}_${month}`, tableRef);

                const filename = `مواضيع_${cls}_${year}-${month.toString().padStart(2, '0')}.xlsx`;
                const filepath = path.join(classPath, filename);
                await wb.xlsx.writeFile(filepath);
                console.log(`Created: ${filepath}`);
            }
        }
    }

    // Full class files
    for (const stage in stageClassMonth) {
        const stagePath = path.join(BASE_DIR, stage);
        for (const cls in stageClassMonth[stage]) {
            const allRecs = [];
            for (const recs of Object.values(stageClassMonth[stage][cls])) {
                allRecs.push(...recs);
            }
            if (allRecs.length === 0) continue;

            allRecs.sort((a, b) => a.dateObj - b.dateObj);
            const data = allRecs.map(r => ({
                التاريخ: r.dateKey,
                الصف: r.class,
                عدد_الحضور: r.count,
                حضور_المرشد: r.teacher,
                عنوان_الموضوع: r.title
            }));

            const wb = new ExcelJS.Workbook();
            const ws = wb.addWorksheet(cls.slice(0, 31));
            ws.columns = [
                { header: 'التاريخ', key: 'التاريخ', width: 15 },
                { header: 'الصف', key: 'الصف', width: 15 },
                { header: 'عدد الحضور', key: 'عدد_الحضور', width: 15 },
                { header: 'حضور المرشد', key: 'حضور_المرشد', width: 20 },
                { header: 'عنوان الموضوع', key: 'عنوان_الموضوع', width: 30 }
            ];
            ws.addRows(data);
            ws.getRow(1).font = { bold: true };
            await autoFitSimple(ws);

            const tableRef = `A1:E${data.length + 1}`;
            addTable(ws, `Table_${cls}_full`, tableRef);

            const fullFilename = `مواضيع_${cls}_الكامل.xlsx`;
            const fullFilepath = path.join(stagePath, fullFilename);
            await wb.xlsx.writeFile(fullFilepath);
            console.log(`Created full class file: ${fullFilepath}`);
        }
    }

    console.log('Session-count attendance processing finished.');
}

// ====================== EXPRESS ENDPOINTS ======================
app.post('/process', async (req, res) => {
    try {
        const { server, token, uid } = req.body;

        const KOBO_SERVER = server || process.env.KOBO_SERVER || "kf.kobotoolbox.org";
        const KOBO_TOKEN = token || process.env.KOBO_API_TOKEN || "caf22d7704035eb531beadd9a130939ff9620879";
        const FORM_UID = uid || process.env.FORM_UID || "a2VR4zHxv4APkh6X7qKvah";

        if (!KOBO_SERVER || !KOBO_TOKEN || !FORM_UID) {
            return res.status(400).json({ error: 'Missing Kobo parameters' });
        }

        console.log('Fetching form...');
        const formJson = await fetchKoboForm(KOBO_SERVER, KOBO_TOKEN, FORM_UID);

        console.log('Fetching submissions...');
        const submissions = await fetchKoboData(KOBO_SERVER, KOBO_TOKEN, FORM_UID);

        console.log(`Fetched ${submissions.length} submissions.`);

        await processAttendancePerStudent(formJson, submissions);
        await processAttendanceSessionCount(formJson, submissions);

        res.json({
            message: 'Processing completed successfully',
            outputs: ['الحضور', 'تفقد المواضيع'],
            submissionCount: submissions.length
        });
    } catch (err) {
        console.error(err);
        res.status(500).json({ error: err.message });
    }
});

app.get('/api/list', (req, res) => {
    const requestedPath = req.query.path || '';

    // Security: prevent directory traversal outside allowed folders
    const normalized = path.normalize('/' + requestedPath).replace(/\\/g, '/');
    const allowedBases = ['/الحضور', '/تفقد المواضيع'];
    let allowed = true;
    for (const base of allowedBases) {
        if (normalized === base || normalized.startsWith(base + '/')) {
            allowed = true;
            break;
        }
    }
    if (!allowed) {
        return res.status(403).json({ error: 'Access denied' });
    }

    const fullPath = path.join(__dirname, normalized);

    fs.readdir(fullPath, { withFileTypes: true }, (err, entries) => {
        if (err) {
            if (err.code === 'ENOENT') {
                return res.status(404).json({ error: 'Directory not found' });
            }
            return res.status(500).json({ error: err.message });
        }

        const filteredEntries = entries.filter(entry => {
            const name = entry.name;
            if (name.startsWith('.') ||
                name === 'node_modules' ||
                name === 'public' ||
                name === 'venv' ||
                name === '__pycache__') {
                return false;
            }
            return true;
        });

        const folders = [];
        const files = [];
        filteredEntries.forEach(entry => {
            if (entry.isDirectory()) {
                folders.push(entry.name);
            } else if (entry.isFile() && entry.name.endsWith('.xlsx')) {
                files.push(entry.name);
            }
        });

        res.json({ folders, files });
    });
});

app.get('/view-excel', async (req, res) => {
    const filePath = req.query.path;
    if (!filePath) {
        return res.status(400).send('Missing path parameter');
    }

    const decodedPath = decodeURIComponent(filePath);
    const normalized = path.normalize('/' + decodedPath).replace(/\\/g, '/');
    const allowedBases = ['/الحضور', '/تفقد المواضيع'];
    let allowed = false;
    for (const base of allowedBases) {
        if (normalized === base || normalized.startsWith(base + '/')) {
            allowed = true;
            break;
        }
    }
    if (!allowed) {
        return res.status(403).send('Access denied');
    }

    const fullPath = path.join(__dirname, decodedPath);
    if (!fs.existsSync(fullPath)) {
        return res.status(404).send('File not found');
    }

    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(fullPath);
        const worksheet = workbook.worksheets[0];

        let html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Excel View</title>';
        html += '<style>';
        html += 'body { font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background: #f5f5f5; }';
        html += '.container { max-width: 1200px; margin: 0 auto; background: white; border-radius: 16px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); padding: 20px; }';
        html += 'h2 { color: #2d3748; border-bottom: 2px solid #48bb78; padding-bottom: 10px; }';
        html += 'table { border-collapse: collapse; width: 100%; margin-top: 20px; }';
        html += 'th { background: #48bb78; color: white; padding: 10px; text-align: center; font-weight: 500; }';
        html += 'td { border: 1px solid #e2e8f0; padding: 8px; text-align: center; }';
        html += 'tr:nth-child(even) { background: #f7fafc; }';
        html += '.absent { background-color: #FFC7CE; color: #9C0006; }';
        html += '.present { background-color: #C6EFCE; color: #006100; }';
        html += '.back-btn { display: inline-block; margin-bottom: 20px; padding: 8px 16px; background: #4299e1; color: white; text-decoration: none; border-radius: 20px; }';
        html += '.back-btn:hover { background: #3182ce; }';
        html += '</style></head><body>';
        html += '<div class="container">';
        html += `<a href="/browser.html" class="back-btn">← Back to Browser</a>`;
        html += `<h2>📄 ${path.basename(decodedPath)}</h2>`;

        html += '<table><thead><tr>';
        const headerRow = worksheet.getRow(1);
        headerRow.eachCell({ includeEmpty: true }, (cell) => {
            html += `<th>${cell.value || ''}</th>`;
        });
        html += '</tr></thead><tbody>';

        for (let i = 2; i <= worksheet.rowCount; i++) {
            const row = worksheet.getRow(i);
            html += '<tr>';
            row.eachCell({ includeEmpty: true }, (cell) => {
                let cellValue = cell.value || '';
                let className = '';
                if (cellValue === 'م') className = 'present';
                if (cellValue === 'غ') className = 'absent';
                html += `<td class="${className}">${cellValue}</td>`;
            });
            html += '</tr>';
        }
        html += '</tbody></table></div></body></html>';
        res.send(html);
    } catch (err) {
        console.error(err);
        res.status(500).send('Error reading Excel file');
    }
});

// Serve static files
app.use(express.static('public'));
app.use('/الحضور', express.static(path.join(__dirname, 'الحضور')));
app.use('/تفقد المواضيع', express.static(path.join(__dirname, 'تفقد المواضيع')));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));