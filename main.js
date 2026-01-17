const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const ExcelProcessor = require('./excel-processor');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true,
            disableHardwareAcceleration: true
        }
    });

    // 启动时最大化窗口（铺满桌面）
    mainWindow.maximize();

    mainWindow.loadFile('index.html');
    // mainWindow.webContents.openDevTools();
}

app.whenReady().then(() => {
    createWindow();

    app.on('activate', function () {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') app.quit();
});

// 处理文件选择请求
ipcMain.handle('select-excel-file', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile'],
        filters: [{
            name: 'Excel Files',
            extensions: ['xlsx', 'xls', 'csv']
        }]
    });

    if (result.canceled || result.filePaths.length === 0) {
        return null;
    }

    const filePath = result.filePaths[0];
    return filePath;
});

let result;
// 处理Excel文件解析
ipcMain.handle('parse-excel-file', async (event, filePath) => {
    try {
        result = await ExcelProcessor.processFile(filePath);
        // console.log("解析结果:", result);
        let newResult = {};
        if (result.success && result.data) {
            newResult = {
                ...result,
                data: {
                    ...result.data,
                    sheets: result.data.sheets.map(sheet => ({
                        ...sheet,
                        studentScores: sheet.studentScores.map(({ "0分题型细分": _, ...resultStudent }) => resultStudent)
                    }))
                }
            }
        }
        return newResult;
    } catch (error) {
        return {
            success: false,
            error: error.message
        };
    }
});

// 处理PDF生成
ipcMain.handle('generate-pdf', async (event, studentData) => {
    const { className, semester, year } = studentData;
    try {
        const fs = require('fs');
        const path = require('path');
        const nunjucks = require('nunjucks');
        const ExcelProcessor = require('./excel-processor');

        // 读取模板文件
        const templatePath = path.join(__dirname, 'demo.html');
        const templateContent = fs.readFileSync(templatePath, 'utf8');

        // 检查是否有解析好的Excel数据
        if (!result || !result.success || !result.data) {
            return {
                success: false,
                error: '请先选择并解析Excel文件'
            };
        }

        const { studentNames, studentsData, sheets, totalSheets, accumulatedZeroScoreTypes } = result.data;

        // 选择文件路径并创建一个以"成绩分析报告"+当前时间为名称的文件夹
        const resultFilePath = await dialog.showOpenDialog(mainWindow, {
            properties: ['openDirectory'],
            title: '选择输出文件夹'
        });

        if (resultFilePath.canceled || resultFilePath.filePaths.length === 0) {
            return null;
        }

        const outputFolderPath = resultFilePath.filePaths[0];
        // console.log("输出文件夹路径:", outputFolderPath);

        // 创建以"成绩分析报告"+当前时间为名称的文件夹
        const reportFolderName = `成绩分析报告${new Date().toLocaleString().replace(/[/:]/g, '-')}`;
        const reportFolderPath = path.join(outputFolderPath, reportFolderName);
        fs.mkdirSync(reportFolderPath, { recursive: true });

        // 使用for...of循环替代forEach以支持async/await
        for (const [name, data] of Object.entries(studentsData)) {
            const examNames = [];
            const studentScores = [];
            const avgScores = [];
            const standardScores = [];
            data.forEach(exam => {
                examNames.push(exam.overallData['考试名称']);
                studentScores.push(exam.studentScore['总分']);
                avgScores.push(exam.overallData['平均分']);
                standardScores.push(exam.studentScore['标准分']);
            });

            // 根据学生处理题型细分
            let questionTypes = [];
            let questionTypeCount = [];
            let questionTypesSecond = [];
            let questionTypeCountSecond = [];
            
            const result = processAccumulatedZeroScoreTypes(accumulatedZeroScoreTypes, name);
            if (result) {
                questionTypes = result.questionTypes || [];
                questionTypeCount = result.questionTypeCount || [];
                questionTypesSecond = result.questionTypesSecond || [];
                questionTypeCountSecond = result.questionTypeCountSecond || [];
            }

            // 计算logo文件的绝对路径
            const path = require('path');
            const logoPath = path.join(__dirname, '卓鸣logo 最终转曲-06.png').replace(/\\/g, '/');
            const watermarkPath = path.join(__dirname, '卓鸣logo 最终转曲-05.png').replace(/\\/g, '/');
            
            const templateData = {
                studentName: name,
                watermarkPath: watermarkPath,
                studentsData: data ? data : {},
                examNames: examNames,
                studentScores: studentScores,
                avgScores: avgScores,
                totalAverageRate: sheets.totalAverageRate,
                questionTypes: questionTypes,
                questionTypeCount: questionTypeCount,
                questionTypesSecond: questionTypesSecond,
                questionTypeCountSecond: questionTypeCountSecond,
                standardScores: standardScores,
                className: className,
                semester: semester,
                year: year,
                logoPath: logoPath
            };

            // 使用nunjucks渲染模板
            const renderedHtml = nunjucks.renderString(templateContent, templateData);

            // 创建临时HTML文件，使用系统临时目录
            const tempDir = os.tmpdir();
            const tempHtmlPath = path.join(tempDir, `temp-report-${name}-${Date.now()}.html`);
            fs.writeFileSync(tempHtmlPath, renderedHtml);

            // 创建新的浏览器窗口来生成PDF
            const { BrowserWindow } = require('electron');
            const pdfWindow = new BrowserWindow({
                show: false,
                webPreferences: {
                    nodeIntegration: true,
                    contextIsolation: false
                }
            });

            // 加载临时HTML文件
            await pdfWindow.loadFile(tempHtmlPath);

            // 等待JavaScript代码执行完成，确保图表渲染完成
            await new Promise(resolve => setTimeout(resolve, 2000));

            // 生成PDF
            const pdfBuffer = await pdfWindow.webContents.printToPDF({
                marginsType: 1,
                pageSize: 'A4',
                printBackground: true,
                printSelectionOnly: false,
                landscape: false,
                scaleFactor: 2.0,
            });

            // 关闭临时窗口
            pdfWindow.close();

            // 删除临时文件
            fs.unlinkSync(tempHtmlPath);

            // 保存PDF文件
            const pdfFilePath = path.join(reportFolderPath, `${name}-成绩分析报告.pdf`);
            fs.writeFileSync(pdfFilePath, pdfBuffer);
        }



        // 返回成功结果，所有学生的PDF报告已生成
        return {
            success: true,
            filePath: reportFolderPath,
            message: `已成功生成 ${Object.keys(studentsData).length} 份成绩分析报告`
        };
    } catch (error) {
        console.error('生成PDF失败:', error);
        return {
            success: false,
            error: error.message
        };
    }
});

const processAccumulatedZeroScoreTypes = (accumulatedZeroScoreTypes, name) => {
    const BOUNDARY = 40;
    let questionTypes = [];
    let questionTypeCount = [];
    let questionTypesSecond = [];
    let questionTypeCountSecond = [];
    if (accumulatedZeroScoreTypes[name]) {
        questionTypes = Object.keys(accumulatedZeroScoreTypes[name]);
        questionTypeCount = Object.values(accumulatedZeroScoreTypes[name]);
    }
    if (questionTypes.length == questionTypeCount.length && questionTypes.length <= BOUNDARY) {
        return { questionTypes, questionTypeCount };
    } else if (questionTypes.length == questionTypeCount.length && questionTypes.length > BOUNDARY) {
        const middleIndex = Math.ceil(questionTypes.length / 2);
        questionTypesSecond = questionTypes.slice(middleIndex);
        questionTypeCountSecond = questionTypeCount.slice(middleIndex);
        questionTypes = questionTypes.slice(0, middleIndex);
        questionTypeCount = questionTypeCount.slice(0, middleIndex);
        return { questionTypes, questionTypeCount, questionTypesSecond, questionTypeCountSecond };

    }
    return;
}
