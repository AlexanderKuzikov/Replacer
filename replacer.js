const AdmZip = require('adm-zip');
const { JSDOM } = require('jsdom');
const fs = require('fs');
const path = require('path');

class DocxCleaner {
    constructor() {
        this.inputDir = './IN';
        this.outputDir = './OUT';
        this.processedFields = 0;
        this.cleanedRuns = 0;
    }

    // Основной метод запуска очистки
    async run() {
        try {
            console.log('=== Утилита очистки DOCX шаблонов ===\n');

            // Проверяем папку IN
            if (!fs.existsSync(this.inputDir)) {
                fs.mkdirSync(this.inputDir, { recursive: true });
                console.log(`Создана папка ${this.inputDir}. Поместите туда DOCX файлы.`);
                return;
            }

            // Ищем DOCX файлы
            const files = fs.readdirSync(this.inputDir)
                .filter(file => file.toLowerCase().endsWith('.docx'));

            if (files.length === 0) {
                console.log('В папке IN не найдено DOCX файлов.');
                return;
            }

            const inputFile = path.join(this.inputDir, files[0]);
            console.log(`Обрабатывается файл: ${files[0]}`);

            // Обрабатываем файл
            await this.processDocx(inputFile);

            console.log('\n=== Очистка завершена ===');
            console.log(`Обработано полей: ${this.processedFields}`);
            console.log(`Удалено разбитых runs: ${this.cleanedRuns}`);
            console.log(`Результат сохранен в папке OUT`);

        } catch (error) {
            console.error('Ошибка:', error.message);
        }
    }

    // Обработка DOCX файла
    async processDocx(inputPath) {
        const zip = new AdmZip(inputPath);
        const zipEntries = zip.getEntries();

        // Ищем document.xml
        let documentEntry = zipEntries.find(entry => 
            entry.entryName === 'word/document.xml' || 
            entry.entryName === 'word/document.xml/'
        );

        if (!documentEntry) {
            throw new Error('Не найден document.xml в DOCX файле');
        }

        // Парсим XML
        const xmlContent = documentEntry.getData().toString('utf8');
        const dom = new JSDOM(xmlContent, { contentType: 'text/xml' });
        const document = dom.window.document;

        // Очищаем структуры полей
        this.cleanFieldStructures(document);

        // Обновляем содержимое в ZIP
        zip.updateFile('word/document.xml', Buffer.from(dom.serialize(), 'utf8'));

        // Сохраняем результат
        this.ensureOutputDir();
        const outputFileName = 're_' + path.basename(inputPath);
        const outputPath = path.join(this.outputDir, outputFileName);
        
        zip.writeZip(outputPath);
        console.log(`Сохранен: ${outputFileName}`);
    }

    // Очистка структур полей
    cleanFieldStructures(document) {
        const allRuns = Array.from(document.getElementsByTagName('w:r'));
        let i = 0;

        while (i < allRuns.length) {
            const run = allRuns[i];
            const fldChar = run.getElementsByTagName('w:fldChar')[0];

            // Если нашли начало поля
            if (fldChar && fldChar.getAttribute('w:fldCharType') === 'begin') {
                const fieldStartIndex = i;
                let fieldStructure = this.extractFieldStructure(allRuns, i);

                if (fieldStructure.isComplete) {
                    this.normalizeFieldStructure(document, fieldStructure);
                    i = fieldStructure.endIndex + 1;
                    this.processedFields++;
                } else {
                    i++;
                }
            } else {
                i++;
            }
        }
    }

    // Извлечение структуры поля
    extractFieldStructure(runs, startIndex) {
        const structure = {
            beginRun: runs[startIndex],
            instrRuns: [],
            separateRun: null,
            textRuns: [],
            endRun: null,
            isComplete: false,
            endIndex: startIndex
        };

        let i = startIndex + 1;
        let phase = 'instr'; // instr -> separate -> text -> end

        while (i < runs.length && phase) {
            const run = runs[i];
            const fldChar = run.getElementsByTagName('w:fldChar')[0];

            if (fldChar) {
                const type = fldChar.getAttribute('w:fldCharType');
                
                if (phase === 'instr' && type === 'separate') {
                    structure.separateRun = run;
                    phase = 'text';
                } else if (phase === 'text' && type === 'end') {
                    structure.endRun = run;
                    structure.endIndex = i;
                    structure.isComplete = true;
                    break;
                }
            } else {
                const instrText = run.getElementsByTagName('w:instrText')[0];
                const text = run.getElementsByTagName('w:t')[0];

                if (phase === 'instr' && instrText) {
                    structure.instrRuns.push(run);
                } else if (phase === 'text' && text) {
                    structure.textRuns.push(run);
                }
            }

            i++;
        }

        return structure;
    }

    // Нормализация структуры поля
    normalizeFieldStructure(document, structure) {
        // Собираем instrText
        if (structure.instrRuns.length > 1) {
            const firstInstr = structure.instrRuns[0].getElementsByTagName('w:instrText')[0];
            let fullInstrText = firstInstr.textContent;

            for (let i = 1; i < structure.instrRuns.length; i++) {
                const instrText = structure.instrRuns[i].getElementsByTagName('w:instrText')[0];
                fullInstrText += instrText.textContent;
                
                // Удаляем run
                structure.instrRuns[i].parentNode.removeChild(structure.instrRuns[i]);
                this.cleanedRuns++;
            }

            firstInstr.textContent = fullInstrText;
        }

        // Собираем текстовое поле
        if (structure.textRuns.length > 1) {
            const firstText = structure.textRuns[0].getElementsByTagName('w:t')[0];
            let fullText = firstText.textContent;

            for (let i = 1; i < structure.textRuns.length; i++) {
                const text = structure.textRuns[i].getElementsByTagName('w:t')[0];
                fullText += text.textContent;
                
                // Удаляем run
                structure.textRuns[i].parentNode.removeChild(structure.textRuns[i]);
                this.cleanedRuns++;
            }

            firstText.textContent = fullText;
        }
    }

    // Создание папки OUT если нужно
    ensureOutputDir() {
        if (!fs.existsSync(this.outputDir)) {
            fs.mkdirSync(this.outputDir, { recursive: true });
        }
    }
}

// Запуск утилиты
const cleaner = new DocxCleaner();
cleaner.run();