const express = require('express');
const multer = require('multer');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const sql = require('mssql');
const path = require('path');
const cors = require('cors');
const { env } = process;
require('dotenv').config({
  path: path.resolve(
      __dirname,
      `./env.${process.env.NODE_ENV ? process.env.NODE_ENV : "test"}`
    ),
});
var corsOptions = {
    origin: '*',
    credentials:true,
    optionsSuccessStatus: 200 // some legacy browsers (IE11, various SmartTVs) choke on 204
  }
const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(cors(corsOptions)).use((req,res,next)=>{
    console.log(req);
    res.setHeader('Access-Control-Allow-Origin',"*");
    next();
});
// 配置上传
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
      cb(null, env.UPLOADS_PATH);
    },
    filename: (req, file, cb) => {
      cb(null, file.originalname);
    }
  });
  
  const upload = multer({ storage: storage });
console.log(env.UPLOADS_PATH);
// MSSQL 连接配置
const config = {
    user: env.USER,
    password: env.PASSWORD,
    server: env.HOST, // You can use 'localhost\\instance' to connect to named instance
    database: env.DATABASE,
    options: {
        encrypt: false, // Use this if you're on Windows Azure
        enableArithAbort: true
    }
};
const monthes = ['一月', '二月', '三月', '四月', '五月', '六月', '七月', '八月', '九月', '十月', '十一月', '十二月'];
const alphabet = Array.from({length: 26}, (_, i) => String.fromCharCode(i + 97).toUpperCase());
const alphabet2 = alphabet.concat(["AA","AB"]);
app.get('/projects', async (req, res) => {
    var _month = req.query.month;
    var _year = req.query.year;
    var maxLength=monthes.length;
    if (_year === undefined || _year === '') {
        _year = new Date().getFullYear();
    }
    if (_month === undefined || _month === '' || _month === '0') {
        _month = 0;
    }else{
        _month = Number(_month)-1
        maxLength = Number(_month)+1;
    }
    console.log(_year,'selected month...',_month)
    try {
        const workbook = new ExcelJS.Workbook();
        for(var i=_month;i<maxLength;i++){
            var emptyValue={};
            var columns=[];
            var columns_row2={ProjName:"项目名称",ParentCode:"项目父编码",ProjGUID:"项目编码",Type:"业态"};
                
            columns.push({ header: monthes[i], key: "month_"+i });
            columns.push({ header: '', key: "month1_"+i });
            columns_row2["month_"+i]="签约(万)";
            columns_row2["month1_"+i]="回款(万)";
            emptyValue["month_"+i]=0;
            emptyValue["month1_"+i]=0;
            
            // 添加一个工作表
            const sheet = workbook.addWorksheet(_year+'年'+(monthes[i]));
            sheet.columns = [
                { header: '项目名称', key: 'ProjName',width:20 },
                { header: '项目父编码', key: 'ParentCode',width:10,hidden:true },
                { header: '项目编码', key: 'ProjGUID',width:10,hidden:true },
                { header: '业态', key: 'Type',width:20 },
            ];
            sheet.columns=sheet.columns.concat(columns);
            
            sheet.addRow(columns_row2);
            var month_index=0;
            for (let i = 4; i < alphabet2.length; i += 2) {
                
                //console.log(`${alphabet2[i]}1:${alphabet2[i+1]}1`);
                sheet.mergeCells(`${alphabet2[i]}1:${alphabet2[i+1]}1`);
                sheet.getCell(`${alphabet2[i]}1`).alignment = { vertical: 'middle',horizontal: 'center' };
                //sheet.getCell(`${alphabet2[i]}1`).value = month[month_index];
                month_index++;
            }
            //console.log(sheet.columns);
            //console.log(emptyValue);
            await sql.connect(config);
            const result = await sql.query`select mb.* from		(select
        
                CASE WHEN ISNULL(p1.ProjName, '') = '' THEN p.ProjName
                     ELSE p1.projname
                END AS 'ProjName' ,
                pm.Type as 'Type',
                p.ParentCode,
				p.ProjGUID 
              from dbo.p_Project as p
              LEFT JOIN p_Project p1 ON p.ParentCode = p1.ProjCode
              LEFT JOIN projectMatcher pm ON p.ProjCode = pm.ProjCode
              where
                (
                  /*北京*/
                  p.BUGUID = '26888FC0-24DC-E311-8D41-D89D672CBEA3' or 
                  /*铜仁*/
                  p.BUGUID = '6EC516A4-659F-E811-8E37-D89D672CBEA3' or 
                  /*苏州*/
                  p.BUGUID = '87BB8FD6-BEB2-E611-B162-D89D672CBEA3'or 
                  /*郑州*/
                  p.BUGUID = '609B0AEE-24DC-E311-8D41-D89D672CBEA3'or 
                  /*大英山*/
                  p.BUGUID = '3DDA941E-25DC-E311-8D41-D89D672CBEA3')
                and p.bgnsaledate > '1900-1-1'
                and p.ProjName not like '测试%'
                and p.ParentCode != 'HNHK.001' AND p.ParentCode != 'HNHK.002'
                and pm.Type is not null
                
                ) mb
                group by mb.Type,mb.ProjName,mb.ParentCode,mb.ProjGUID
                ORDER BY mb.ProjName`;
            var count=2;
            var startRow=2;
            let currentContent = null;
            const columnKey = 'A';
            const columnKeyB = 'B';
            const columnKeyC = 'C';
            const columnKeyD = 'D';
            sheet.mergeCells(`${columnKey}${1}:${columnKey}${2}`);
            sheet.getCell(`${columnKey}${1}`).alignment = { vertical: 'middle',horizontal: 'center' };
            sheet.mergeCells(`${columnKeyB}${1}:${columnKeyB}${2}`);
            sheet.getCell(`${columnKeyB}${1}`).alignment = { vertical: 'middle',horizontal: 'center' };
            sheet.mergeCells(`${columnKeyC}${1}:${columnKeyC}${2}`);
            sheet.getCell(`${columnKeyC}${1}`).alignment = { vertical: 'middle',horizontal: 'center' };
            sheet.mergeCells(`${columnKeyD}${1}:${columnKeyD}${2}`);
            sheet.getCell(`${columnKeyD}${1}`).alignment = { vertical: 'middle',horizontal: 'center' };
            result.recordset.forEach(row => {
                //console.log(row); // 打印每一行数据
                count++;
                var data={...{ ProjName: row.ProjName, ParentCode: row.ParentCode, ProjGUID: row.ProjGUID, Type: row.Type },...emptyValue}
                //console.log(data);
                sheet.addRow(data);
                sheet.getCell(`${columnKey}${startRow}`).alignment = { vertical: 'middle',horizontal: 'center' };
                if (row.ProjName !== currentContent) {
                    if (startRow && (count - startRow > 1)) { // 有超过一个单元格的内容相同
                        sheet.mergeCells(`${columnKey}${startRow}:${columnKey}${count - 1}`);
                        
                    }
                    startRow = count;
                    currentContent = row.ProjName;
                }
                
                
            });
            if (startRow && (count - startRow > 1)) { // 有超过一个单元格的内容相同
                sheet.mergeCells(`${columnKey}${startRow}:${columnKey}${count}`);
                sheet.getCell(`${columnKey}${startRow}`).alignment = { vertical: 'middle',horizontal: 'center' };
            }
        }
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename="template.xlsx"');

        // 将Excel文件写入到Response中
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        console.error(err);
        res.status(500).send('Error getting projects');
    } finally {
        sql.close();
    }
})
// 路由处理文件上传和数据插入
app.post('/upload', upload.single('excel'), async (req, res) => {
    let pool;
    try {
            pool = await new sql.ConnectionPool(config)
            //await sql.connect(config);
            console.log(req.file);
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.readFile(req.file.path);
            let tableRows = [];
            await new Promise((s_resolve,s_reject)=>{
                var count=0;
                workbook.eachSheet(async (sheet, id) => {
                    const expectedDate=chineseDateToMSSQLDate(sheet.name);
                    console.log(expectedDate);
                     // Array to store HTML table rows
                     var rowCount=0;
                    await new Promise((resolve,reject)=>{
                        sheet.eachRow({ includeEmpty: false }, async (row, rowNumber) => {
                            if (rowNumber > 2) { // Assuming first row is headers
                                const checkQuery = `
                                    IF EXISTS (
                                        SELECT 1 FROM expected 
                                        WHERE ProjGUID = @ProjGUID AND ExpectedDate = @ExpectedDate
                                    )
                                    BEGIN
                                        -- Update the existing record
                                        UPDATE expected 
                                        SET ParentCode = @ParentCode, ProjShortName = @ProjShortName, ExpectedAmount = @ExpectedAmount, ExpectedRepay = @ExpectedRepay
                                        WHERE ProjGUID = @ProjGUID AND ExpectedDate = @ExpectedDate
                                    END
                                    ELSE
                                    BEGIN
                                        -- Insert a new record
                                        INSERT INTO expected (ProjGUID, ParentCode, ProjShortName, ExpectedDate, ExpectedAmount, ExpectedRepay) 
                                        VALUES (@ProjGUID, @ParentCode, @ProjShortName, @ExpectedDate, @ExpectedAmount, @ExpectedRepay)
                                    END
                                `;
                                await pool.connect()
                                    const request = pool.request();
                                    request.input('ProjGUID', sql.UniqueIdentifier, row.getCell(3).value);
                                    request.input('ParentCode', sql.VarChar, row.getCell(2).value);
                                    request.input('ProjShortName', sql.VarChar, row.getCell(4).value);
                                    request.input('ExpectedDate', sql.DateTime, expectedDate);
                                    request.input('ExpectedAmount', sql.Money, row.getCell(5).value);
                                    request.input('ExpectedRepay', sql.Money, row.getCell(6).value);
                                    try{
                                        await request.query(checkQuery);
                                        console.log(row.getCell(2).value,row.getCell(3).value,expectedDate,row.getCell(4).value,row.getCell(5).value,row.getCell(6).value)
                                        tableRows.push(`
                                                <tr>
                                                    <td>${row.getCell(1).value}</td>
                                                    <td>${row.getCell(2).value}</td>
                                                    <td>${row.getCell(3).value}</td>
                                                    <td>${row.getCell(4).value}</td>
                                                    <td>${sheet.name}</td>
                                                    <td>${row.getCell(5).value}</td>
                                                    <td>${row.getCell(6).value}</td>
                                                </tr>
                                            `);
                                    }catch(err){
                                        console.error(err);
                                        tableRows.push(`
                                                <tr>
                                                    <td colspan="7">Error processing file at row ${rowCount}, ${err.message}</td>
                                                </tr>
                                            `);
                                        //res.status(500).send('Error processing file at row '+rowCount);
                                    }
                                    
                            
                            
                            }
                            rowCount++;
                        });
                        var interval=setInterval(() => {
                            if(rowCount==sheet.rowCount){
                                clearInterval(interval)
                                count++;
                                resolve(tableRows);
                            }
                        }, 100);
                    });
                });
                var s_interval=setInterval(() => {
                    if(count==workbook.worksheets.length ){
                        clearInterval(s_interval)
                        s_resolve(tableRows);
                    }
                }, 100);
            });
            console.log('done')
            const htmlTable = `
                <table class="import_table" border="1">
                    <thead>
                        <tr>
                            <th>项目名称</th>
                            <th>项目父编码</th>
                            <th>项目编码</th>
                            <th>业态</th>
                            <th>月份</th>
                            <th>签约(万)</th>
                            <th>回款(万)</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${tableRows.join('')}
                    </tbody>
                </table>
            `;
        res.send(htmlTable);
    } catch (err) {
        console.error(err);
        res.status(500).send('Error processing file');
    } finally {
        if (pool) {
            await pool.close();
        }
    }
});
function chineseDateToMSSQLDate(input) {
    // Define a map for the Chinese months
    const monthMap = {
        '一': '01', '二': '02', '三': '03', '四': '04',
        '五': '05', '六': '06', '七': '07', '八': '08',
        '九': '09', '十': '10', '十一': '11', '十二': '12'
    };

    // Regular expression to extract the year and month part
    const regex = /(\d{4})年(.+)月/;
    const matches = input.match(regex);

    if (matches && matches.length === 3) {
        const year = matches[1];
        var month = monthMap[matches[2]]; // Get the month number from the map
        
        
        if(month===undefined){
            try{
                month=Number(matches[2])<10?"0"+matches[2]:matches[2];
            }catch(err){
                throw new Error("Invalid month "+err.message);
            }
            
        }
        console.log(matches[2],month);
        if (month) {
            // Format into a date string that MSSQL can accept (first day of the specified month)
            return `${year}-${month}-01`;
        } else {
            throw new Error("Invalid month");
        }
    } else {
        throw new Error("Invalid date format");
    }
}
// 启动服务器
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT},${process.env.NODE_ENV}`);
});

// Create table expected (
// 	ProjGUID uniqueidentifier NOT NULL,
// 	ParentCode varchar(40),
// 	ProjShortName varchar(40), 
// 	ExpectedDate datetime,
// 	ExpectedAmount money,
// 	ExpectedRepay money
// 	primary key (ProjGUID,ExpectedDate,ProjShortName)
// );