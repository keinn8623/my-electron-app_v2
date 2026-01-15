const XLSX = require('xlsx');

class ExcelProcessor {
    
    static async processFile(filePath) {
        try {
            // 使用xlsx读取文件，自动处理编码
            const workbook = XLSX.readFile(filePath, { cellText: true, cellDates: true });
            
            // 处理所有sheet
            const sheetsData = [];
            
            let totalAverageRate = {};
            // 遍历所有sheet
            for (const sheetName of workbook.SheetNames) {
                if (sheetName === '题型分类') {
                    break;
                }
                const worksheet = workbook.Sheets[sheetName];
                
                // 将工作表转换为JSON格式，设置默认值为空字符串
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
                    defval: '',
                    header: 1 // 使用数组格式，第一行作为索引0
                });
                
                // 检查数据行数是否足够
                if (jsonData.length < 5) {
                    // 跳过数据不足的sheet，继续处理其他sheet
                    continue;
                }
                
                // 第2行（数组索引1）：题型
                const questionTypes = jsonData[1];
                // 第3行（数组索引2）：题型细分
                const detailTypes = jsonData[2];
                // 第4行（数组索引3）：满分分数
                const fullScores = jsonData[3].map(score => parseFloat(score) || 0);
                // 第5行（数组索引4）开始：学生数据
                const studentData = jsonData.slice(4);
                
                // 按学生计算每个题型的分数
                let studentScores = [];

                // 根据fullScores,把相同题型的分数累加起来
                const questionTypeTotals = {};
                fullScores.forEach((score, index) => {
                    const type = questionTypes[index];
                    if (type && type.trim()) {
                        const trimmedType = type.trim();
                        questionTypeTotals[trimmedType] = (questionTypeTotals[trimmedType] || 0) + score;
                    }
                });
                
                // 遍历每个学生数据
                studentData.forEach((row, studentIndex) => {
                    // 跳过空行
                    const isEmptyRow = row.every(cell => cell === '' || cell === null || cell === undefined);
                    if (isEmptyRow) {
                        return;
                    }
                    
                    // 初始化学生分数对象
                    const studentScore = {
                        // "学生序号": studentIndex + 1,
                        "学生姓名": row[0] || '',
                        "老师": row[1] || '',
                        "校区": row[2] || ''
                    };
                    
                    // 统计0分题对应的题型细分
                    const zeroScoreDetailTypes = {};
                    
                    // 按题型计算分数
                    const questionTypeTotals = {};
                    
                    // 初始化题型分数
                    questionTypes.forEach((type, index) => {
                        if (type && type.trim()) {
                            const trimmedType = type.trim();
                            questionTypeTotals[trimmedType] = 0;
                        }
                    });
                    
                    // 计算每个题型的分数，并统计0分题的题型细分
                    row.forEach((score, index) => {
                        const type = questionTypes[index];
                        if (type && type.trim()) {
                            const trimmedType = type.trim();
                            const scoreValue = parseFloat(score) || 0;
                            questionTypeTotals[trimmedType] += scoreValue;
                        }
                        
                        // 统计0分题的题型细分（跳过前3列学生信息和最后2列总分排名）
                        if (index >= 3 && index < row.length - 2) {
                            const scoreValue = parseFloat(score) || 0;
                            if (scoreValue === 0) {
                                const detailType = detailTypes[index] || '';
                                if (detailType && detailType.trim()) {
                                    const trimmedDetailType = detailType.trim();
                                    zeroScoreDetailTypes[trimmedDetailType] = (zeroScoreDetailTypes[trimmedDetailType] || 0) + 1;
                                }
                            }
                        }
                    });

                    // 合并到学生分数对象
                    Object.assign(studentScore, questionTypeTotals);
                    
                    // 添加0分题题型细分统计
                    studentScore["0分题型细分"] = zeroScoreDetailTypes;
                    
                    // 添加总分和排名（如果有）
                    studentScore["总分"] = row[row.length - 2] || '';
                    studentScore["排名"] = parseInt(row[row.length - 1] || '') || 0;
                    studentScores.push(studentScore);
                });

                let totalCalcRate = [];
                let totalNumTheoryRate = [];
                let totalCountRate = [];
                let totalGeoRate = [];
                let totalAppRate = [];
                let totalTripRate = [];
                let totalCombinedRate = [];

                // 计算每个学生的模块正确率
                studentScores.forEach(studentScore => {
                    if(typeof questionTypeTotals['计算'] !== 'number') {
                        studentScore["计算正确率"] = '-'
                    } else {
                        const calcRate = parseFloat((studentScore['计算'] / questionTypeTotals['计算'] * 100).toFixed(2));
                        studentScore["计算正确率"] = calcRate.toString() + '%';
                        totalCalcRate.push(calcRate);
                    }
                    
                    if(typeof questionTypeTotals['数论'] !== 'number') {
                        studentScore["数论正确率"] = '-'
                    } else {
                        const numTheoryRate = parseFloat((studentScore['数论'] / questionTypeTotals['数论'] * 100).toFixed(2));
                        studentScore["数论正确率"] = numTheoryRate.toString() + '%';
                        totalNumTheoryRate.push(numTheoryRate);
                    }
                
                    if(typeof questionTypeTotals['应用'] !== 'number') {
                        studentScore["应用正确率"] = '-'
                    } else {
                        const appRate = parseFloat((studentScore['应用'] / questionTypeTotals['应用'] * 100).toFixed(2));
                        studentScore["应用正确率"] = appRate.toString() + '%';
                        totalAppRate.push(appRate);
                    }
                     
                    if(typeof questionTypeTotals['组合'] !== 'number') {
                        studentScore["组合正确率"] = '-'
                    } else {
                        const combinedRate = parseFloat((studentScore['组合'] / questionTypeTotals['组合'] * 100).toFixed(2));
                        studentScore["组合正确率"] = combinedRate.toString() + '%';
                        totalCombinedRate.push(combinedRate);
                    }

                    
                    if(typeof questionTypeTotals['计数'] !== 'number') {
                        studentScore["计数正确率"] = '-'
                    } else {
                        const countRate = parseFloat((studentScore['计数'] / questionTypeTotals['计数'] * 100).toFixed(2));
                        studentScore["计数正确率"] = countRate.toString() + '%';
                        totalCountRate.push(countRate);
                    }

                    if(typeof questionTypeTotals['行程'] !== 'number') {
                        studentScore["行程正确率"] = '-'
                    } else {
                        const tripRate = parseFloat((studentScore['行程'] / questionTypeTotals['行程'] * 100).toFixed(2));
                        studentScore["行程正确率"] = tripRate.toString() + '%';
                        totalTripRate.push(tripRate);
                    }

                    if(typeof questionTypeTotals['几何'] !== 'number') {
                        studentScore["几何正确率"] = '-'
                    } else {
                        const geoRate = parseFloat((studentScore['几何'] / questionTypeTotals['几何'] * 100).toFixed(2));
                        studentScore["几何正确率"] = geoRate.toString() + '%';
                        totalGeoRate.push(geoRate);
                    }
                })

                // 计算总体平均正确率
                totalAverageRate = {
                    '计算正确率': totalCalcRate.length > 0 ? parseFloat(totalCalcRate.reduce((acc, cur) => acc + cur, 0) / totalCalcRate.length).toFixed(2) : '0',
                    '数论正确率': totalNumTheoryRate.length > 0 ? parseFloat(totalNumTheoryRate.reduce((acc, cur) => acc + cur, 0) / totalNumTheoryRate.length).toFixed(2) : '0',
                    '应用正确率': totalAppRate.length > 0 ? parseFloat(totalAppRate.reduce((acc, cur) => acc + cur, 0) / totalAppRate.length).toFixed(2) : '0',
                    '组合正确率': totalCombinedRate.length > 0 ? parseFloat(totalCombinedRate.reduce((acc, cur) => acc + cur, 0) / totalCombinedRate.length).toFixed(2) : '0',
                    '计数正确率': totalCountRate.length > 0 ? parseFloat(totalCountRate.reduce((acc, cur) => acc + cur, 0) / totalCountRate.length).toFixed(2) : '0',
                    '行程正确率': totalTripRate.length > 0 ? parseFloat(totalTripRate.reduce((acc, cur) => acc + cur, 0) / totalTripRate.length).toFixed(2) : '0',
                    '几何正确率': totalGeoRate.length > 0 ? parseFloat(totalGeoRate.reduce((acc, cur) => acc + cur, 0) / totalGeoRate.length).toFixed(2) : '0',
                }

                // 按照排名从小到大排序
                studentScores.sort((a, b) => a.排名 - b.排名);

                // overall 数据
                const overAllData = [];
                let totalScore = 0;
                let averageScore = 0;
                let middleScore = 0;
                let highestScore = 0;
                let lowestScore = 0;
                let standardDeviation = 0;
                let sumOfSquaredDifferences = 0;

                // 计算总分平均分
                studentScores.forEach(studentScore => {
                    totalScore += parseFloat(studentScore["总分"]) || 0;
                })
                averageScore = parseFloat((totalScore / studentScores.length).toFixed(2));

                // 计算中间分
                if (studentScores.length % 2 === 0) {  
                    // 偶数个学生，取中间两个学生的平均分
                    middleScore = parseFloat((
                        (parseFloat(studentScores[studentScores.length / 2 - 1]["总分"]) || 0) +
                        (parseFloat(studentScores[studentScores.length / 2]["总分"]) || 0)
                    ) / 2).toFixed(2);
                } else { 
                    // 奇数个学生，取中间那个学生的分数
                    middleScore = parseFloat((parseFloat(studentScores[Math.floor(studentScores.length / 2)]["总分"]) || 0)).toFixed(2);
                }

                // 计算最高分和最低分
                highestScore = parseFloat(studentScores[0]["总分"]).toFixed(2) || 0;
                lowestScore = parseFloat(studentScores[studentScores.length - 1]["总分"]).toFixed(2) || 0;

                // 计算标准差（标准差越大，学生成绩的分布越分散）
                studentScores.forEach(studentScore => {
                    const score = parseFloat(studentScore["总分"]) || 0;
                    sumOfSquaredDifferences += Math.pow(score - averageScore, 2);
                });
                standardDeviation = parseFloat(Math.sqrt(sumOfSquaredDifferences / studentScores.length).toFixed(2));

                /* 计算每个学生的标准分
                - 标准分可以反映学生成绩在整体中的相对位置
                - 标准分为正表示高于平均分，为负表示低于平均分
                - 标准分的绝对值越大，说明学生成绩与平均分的差距越大
                - 标准分可以用于不同考试之间的成绩比较
                */
                studentScores.forEach(studentScore => {
                    const score = parseFloat(studentScore["总分"]) || 0;
                    // 标准分 = (学生分数 - 平均分) / 标准差
                    let standardScore = 0;
                    if (standardDeviation !== 0) {
                        standardScore = parseFloat(((score - averageScore) / standardDeviation).toFixed(2));
                    }
                    // 添加标准分到学生分数对象
                    studentScore["标准分"] = standardScore;
                });

                overAllData.push({
                    "考试名称": sheetName,
                    "学生人数": studentScores.length,
                    "试卷总分": worksheet['AD3'] ? worksheet['AD3'].v : '',  // 读取AD3单元格的内容
                    "平均分": averageScore,
                    "中间分": middleScore,
                    "最高分": highestScore,
                    "最低分": lowestScore,
                    "标准差": standardDeviation
                })
                
                // 各模块数据
                const moduleScores = [];
                // 计算模块平均分
                let calcAverageScore = 0;
                let calcTotalScore = 0;
                studentScores.forEach(studentScore => {
                    calcTotalScore += parseFloat(studentScore["计算"]) || 0;
                })
                calcAverageScore = parseFloat((calcTotalScore / studentScores.length).toFixed(2));

                // 计数模块平均分
                let countAverageScore = 0;
                let countTotalScore = 0;
                studentScores.forEach(studentScore => {
                    countTotalScore += parseFloat(studentScore["计数"]) || 0;
                })
                countAverageScore = parseFloat((countTotalScore / studentScores.length).toFixed(2));

                // 数论模块平均分
                let numTheoryAverageScore = 0;
                let numTheoryTotalScore = 0;
                studentScores.forEach(studentScore => {
                    numTheoryTotalScore += parseFloat(studentScore["数论"]) || 0;
                })
                numTheoryAverageScore = parseFloat((numTheoryTotalScore / studentScores.length).toFixed(2));

                // 几何模块平均分
                let geoTheoryAverageScore = 0;
                let geoTheoryTotalScore = 0;
                studentScores.forEach(studentScore => {
                    geoTheoryTotalScore += parseFloat(studentScore["几何"]) || 0;
                })
                geoTheoryAverageScore = parseFloat((geoTheoryTotalScore / studentScores.length).toFixed(2));

                // 应用模块平均分
                let appAverageScore = 0;
                let appTotalScore = 0;
                studentScores.forEach(studentScore => {
                    appTotalScore += parseFloat(studentScore["应用"]) || 0;
                })
                appAverageScore = parseFloat((appTotalScore / studentScores.length).toFixed(2));

                // 行程模块平均分
                let tripAverageScore = 0;
                let tripTotalScore = 0;
                studentScores.forEach(studentScore => {
                    tripTotalScore += parseFloat(studentScore["行程"]) || 0;
                })
                tripAverageScore = parseFloat((tripTotalScore / studentScores.length).toFixed(2));

                // 组合模块平均分
                let combinedAverageScore = 0;
                let combinedTotalScore = 0;
                studentScores.forEach(studentScore => {
                    combinedTotalScore += parseFloat(studentScore["总分"]) || 0;
                })
                combinedAverageScore = parseFloat((combinedTotalScore / studentScores.length).toFixed(2));

                moduleScores.push({
                    "计算平均分": calcAverageScore,
                    "计数平均分": countAverageScore,
                    "数论平均分": numTheoryAverageScore,
                    "几何平均分": geoTheoryAverageScore,
                    "应用平均分": appAverageScore,
                    "行程平均分": tripAverageScore,
                    "组合平均分": combinedAverageScore,
                })

                
                // 按题型汇总所有学生的分数
                const questionTypeScores = {};
                
                // 初始化题型汇总
                questionTypes.forEach((type, index) => {
                    if (type && type.trim()) {
                        const trimmedType = type.trim();
                        questionTypeScores[trimmedType] = {
                            fullScore: fullScores[index],
                            totalScore: 0,
                            studentCount: studentScores.length,
                            averageScore: 0
                        };
                    }
                });
                
                // 计算每个题型的总分数
                studentScores.forEach(studentScore => {
                    Object.entries(studentScore).forEach(([key, value]) => {
                        if (questionTypeScores[key]) {
                            questionTypeScores[key].totalScore += parseFloat(value) || 0;
                        }
                    });
                });
                
                // 计算平均分
                Object.values(questionTypeScores).forEach(score => {
                    if (score.studentCount > 0) {
                        score.averageScore = parseFloat((score.totalScore / score.studentCount).toFixed(2));
                    }
                });
                
                // 转换为数组格式
                const questionTypeScoresArray = Object.entries(questionTypeScores).map(([type, data]) => ({
                    type: type,
                    fullScore: data.fullScore,
                    totalScore: data.totalScore,
                    studentCount: data.studentCount,
                    averageScore: data.averageScore
                }));

                let groupByStudentScores = [];
                groupByStudentScores = studentScores.map(studentScore => ({
                    "学生姓名": studentScore["学生姓名"],
                    "老师": studentScore["老师"],
                    "校区": studentScore["校区"],
                    "数论": !studentScore["数论"] && studentScore["数论"] !== 0? "-": studentScore["数论"],
                    "数论正确率": studentScore["数论正确率"],
                    "计数": !studentScore["计数"] && studentScore["计数"] !== 0? "-": studentScore["计数"],
                    "计数正确率": studentScore["计数正确率"],
                    "组合": !studentScore["组合"] && studentScore["组合"] !== 0? "-": studentScore["组合"],
                    "组合正确率": studentScore["组合正确率"],
                    "几何": !studentScore["几何"] && studentScore["几何"] !== 0? "-": studentScore["几何"],
                    "几何正确率": studentScore["几何正确率"],
                    "行程": !studentScore["行程"] && studentScore["行程"] !== 0? "-": studentScore["行程"],
                    "行程正确率": studentScore["行程正确率"],
                    "应用": !studentScore["应用"] && studentScore["应用"] !== 0? "-": studentScore["应用"],
                    "应用正确率": studentScore["应用正确率"],
                    "计算": !studentScore["计算"] && studentScore["计算"] !== 0? "-": studentScore["计算"],
                    "计算正确率": studentScore["计算正确率"],
                    "排名": studentScore['排名'],
                    "总分": studentScore['总分'],
                    "标准分": studentScore['标准分'],
                    "0分题型细分": studentScore["0分题型细分"] || {},
                }));

                
                // 添加当前sheet的数据到结果数组
                sheetsData.push({
                    sheetName: sheetName,
                    studentScores: groupByStudentScores,
                    questionTypeScores: questionTypeScoresArray,
                    studentCount: studentScores.length,
                    overallData: overAllData[0], // 每个sheet只有一个overall数据对象
                });
                
            }
            
            // 检查是否有有效数据
            if (sheetsData.length === 0) {
                return {
                    success: false,
                    error: 'Excel文件中没有符合要求的工作表（每个工作表至少需要4行数据）'
                };
            }
            
            // 收集所有学生姓名并按学生分组
            const studentNames = new Set();
            const studentsData = {};
            const accumulatedZeroScoreTypes = {};
            
            // 遍历所有sheet，收集学生数据并累加0分题型细分
            sheetsData.forEach(sheet => {
                sheet.studentScores.forEach(student => {
                    const studentName = student["学生姓名"];
                    if (studentName && studentName.trim()) {
                        studentNames.add(studentName);
                        
                        // 按学生姓名分组，存储该学生在所有sheet中的数据
                        if (!studentsData[studentName]) {
                            studentsData[studentName] = [];
                            accumulatedZeroScoreTypes[studentName] = {};
                        }
                        
                        // 添加该学生在当前sheet中的数据
                        studentsData[studentName].push({
                            sheetName: sheet.sheetName,
                            studentScore: student,
                            overallData: sheet.overallData
                        });
                        
                        // 累加0分题型细分
                        const zeroScoreTypes = student["0分题型细分"] || {};
                        Object.entries(zeroScoreTypes).forEach(([detailType, count]) => {
                            accumulatedZeroScoreTypes[studentName][detailType] = 
                                (accumulatedZeroScoreTypes[studentName][detailType] || 0) + count;
                        });
                    }
                });
            });

            // 将Set转换为数组
            const uniqueStudentNames = Array.from(studentNames).sort();
            
            // 添加累加后的0分题型细分到返回数据中
            const resultData = {
                sheets: sheetsData,
                totalSheets: sheetsData.length,
                studentNames: uniqueStudentNames,
                studentsData: studentsData,
                accumulatedZeroScoreTypes: accumulatedZeroScoreTypes,
                totalAverageRate: totalAverageRate,
            };
            // console.log('accumulatedZeroScoreTypes:',resultData.accumulatedZeroScoreTypes)
            
            return {
                success: true,
                data: resultData
            };
        } catch (error) {
            return {
                success: false,
                error: error.message
            };
        }
    }
}

module.exports = ExcelProcessor;
