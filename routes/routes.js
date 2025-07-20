const express = require('express');
const partsModel=require("../models/partsModel.js");
const router=express.Router();
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const QuickChart  = require('quickchart-js');
const axios       = require('axios');


// 2) lists of subtypes (in the exact order you need columns)
const castingList = [
  {code:'C29',label:'Dent Mark / Contusion'},{code:'C19',label:'Dirt Inclusion'},{code:'C13',label:'Rough Surface'},{code:'C13',label:'Kajari'},
  {code:'C6',label:'Blowhole'},{code:'C8',label:'Oxide Inclusions'},{code:'C90',label:'Over Grinding'},{code:'C5',label:'Shrinkage'},{code:'C31',label:'Pin Hole'},
  {code:'C11',label:'Cold Lap'},{code:'C2',label:'Crack'},{code:'C14',label:'Diecoat Peel-Off'},{code:'C98',label:'Unmachined'},{code:'C99',label:'Casting Other'},{code:'C6',label:'Flange'}
];
const machiningList = [
  {code:'M16',label:'Machining dent hub'},{code:'M16',label:'Bore'},{code:'M14',label:'Un Machined'},{code:'M11',label:'Broken Chips'},{code:'M10',label:'Chamfer Error'},{code:'M90',label:'Over Grinding'},
  {code:'M99',label:'M/C Others'},{code:'M15',label:'Machining Mark'},{code:'M',label:'Deburring Tool Mark'},{code:'M',label:'Burr'},
  {code:'M',label:'Chip Sticking'},{code:'M',label:'Scratch'},{code:'M',label:'Disk Dent'},{code:'M',label:'Step'}
];

// helper to turn 1→A, 27→AA, etc.
function colLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

router.get('/getPartsData',async (req,res)=>{
 let scannedPartsData=await partsModel.find();
 const filter=/^c/
 const defects=await partsModel.aggregate([
  {
    $match: {
      defecttype: { $regex: filter },
    },
  },
  {
    $group: {
      _id: '$defecttype',
      defectQty: { $sum: 1 },
    },
  },
  {
    $sort: { defectQty: -1 },
  },
  {
    $group: {
      _id: null,
      defects: { $push: '$$ROOT' },
      totalQty: { $sum: '$defectQty' },
    },
  },
  {
    $project: {
      _id: 0,
      defects: {
        $map: {
          input: { $range: [0, { $size: '$defects' }] },
          as: 'index',
          in: {
            srNo: { $add: ['$$index', 1] },
            defectName: { $arrayElemAt: ['$defects._id', '$$index'] },
            defectQty: { $arrayElemAt: ['$defects.defectQty', '$$index'] },
            cumQty: {
              $let: {
                vars: {
                  sortedQty: { $reduce: {
                    input: { $slice: ['$defects.defectQty', 0, { $add: ['$$index', 1] }] },
                    initialValue: 0,
                    in: { $add: ['$$value', '$$this'] },
                  }},
                },
                in: '$$sortedQty',
              },
            },
            cummPercent: {
              $cond: [
                { $eq: ['$totalQty', 0] },
                0,
                { $multiply: [{ $divide: [
                  { $let: {
                    vars: {
                      sortedQty: { $reduce: {
                        input: { $slice: ['$defects.defectQty', 0, { $add: ['$$index', 1] }] },
                        initialValue: 0,
                        in: { $add: ['$$value', '$$this'] },
                      }},
                    },
                    in: '$$sortedQty',
                  }},
                  '$totalQty',
                ]}, 100] },
              ],
            },
          },
        },
      },
      totalQty: 1,
    },
  },
]);
console.log(defects)
 res.send(scannedPartsData)
})

router.post("/api/addPartsData",async (req,res)=>{
    let {model,defect,modelState, subDefect}=req.body;
    console.log(modelState)
    let ok,rework,reject;
    if(Number(modelState)===0){
      ok=0
      rework=1
      reject=0
    }
    else if(Number(modelState)===1){
      ok=1
      rework=0
      reject=0
    }
    else{
      ok=0
      rework=0
      reject=1
    }
    await partsModel.create(
                {                  
                    model,
                    customer:'msil',
                    defectType:defect,
                    defectSubType:subDefect,
                    ok,
                    rework,
                    date_time:String(new Date()),
                    reject
                }
            )
    res.sendStatus(200)
}
)

router.get("/api/generate-inspection-report",async(req,res)=>{
try{

  const groupStage = {
    _id: { model: '$model', customer: '$customer' },
    ok:        { $sum: '$ok' },
    rework:    { $sum: '$rework' },
    reject:    { $sum: '$reject' },
    totalChecked: {
      $sum: { $add: ['$ok', '$rework', '$reject'] }
    }
  };

  // add one accumulator per casting subtype
  castingList.forEach(item => {
    const key = item.label.replace(/\W/g,'_'); // e.g. "Pin Hole"→"Pin_Hole"
    groupStage[`cast_${key}`] = {
      $sum: {
        $cond: [
          { $and: [
              { $eq: ['$defectType','casting'] },
              { $eq: ['$defectSubType', item.label] }
            ]},
          { 
      $add: [
        { $ifNull: ['$rework', 0] },
        { $ifNull: ['$reject', 0] }
      ]
    },
          0
        ]
      }
    };
  });

  // ... and for machining
  machiningList.forEach(item => {
    const key = item.label.replace(/\W/g,'_');
    groupStage[`mach_${key}`] = {
      $sum: {
        $cond: [
          { $and: [
              { $eq: ['$defectType','machining'] },
              { $eq: ['$defectSubType', item.label] }
            ]},
         { 
      $add: [
        { $ifNull: ['$rework', 0] },
        { $ifNull: ['$reject', 0] }
      ]
    },
          0
        ]
      }
    };
  });

  // 1) run the aggregation
  const agg = await partsModel.aggregate([
    { $group: groupStage },
    { $sort: { '_id.model': 1, '_id.customer': 1 } }
  ]);

  // 2) Massage into flat rows
  const rows = agg.map(doc => {
    const base = {
      model: doc._id.model,
      customer: doc._id.customer,
      totalChecked: doc.totalChecked,
      ok: doc.ok,
      rework: doc.rework,
      reject: doc.reject
    };
    // attach each casting col in order:
    castingList.forEach(item => {
      base[item.label] = doc[`cast_${item.label.replace(/\W/g,'_')}`] || 0;
    });
    // attach machining
    machiningList.forEach(item => {
      base[item.label] = doc[`mach_${item.label.replace(/\W/g,'_')}`] || 0;
    });
    return base;
  });

  // --- B) Build the Excel with ExcelJS ---
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Incoming Inspection');

  // define where each block starts/ends
  const COL = {
    MODEL:       1,
    CUSTOMER:    2,
    TOTAL:       3,
    CAST_START:  4,
    CAST_END:    4 + castingList.length - 1,
    MACH_START:  null,  // fill next
    OK:          null,  // fill after machining
    RW_PCT:      null,
    CP_PCT:      null,
    MP_PCT:      null
  };
  COL.MACH_START = COL.CAST_END + 1;
  COL.MACH_END   = COL.MACH_START + machiningList.length - 1;
  COL.TOTAL_REPROCESS = COL.MACH_END + 1;
  COL.OK         = COL.TOTAL_REPROCESS + 1;
  COL.RW_PCT     = COL.OK + 1;
  COL.CP_PCT     = COL.RW_PCT + 1;
  COL.MP_PCT     = COL.CP_PCT + 1;
  const LAST_COL = COL.MP_PCT;

  // 1) Title row
  ws.mergeCells(`A1:${colLetter(LAST_COL)}1`);
  ws.getCell('A1').value = `20-09-24 To 20-09-24   INCOMING INSPECTION REPORT`;
  ws.getCell('A1').font = { size:14, bold:true };
  ws.getCell('A1').alignment = { horizontal:'center' };
  ws.getRow(1).height = 24;

  // 2) Header rows 2 & 3
  ['A','B','C'].forEach((col, idx) => {
    ws.mergeCells(`${col}2:${col}3`);
    const txt = idx===0? 'MODEL' : idx===1? 'CUSTOMER' : 'TOTAL CHECKED';
    ws.getCell(`${col}2`).value = txt;
    ws.getCell(`${col}2`).font = { bold:true };
    ws.getCell(`${col}2`).alignment = { vertical:'middle', horizontal:'center' };
    ws.getColumn(idx+1).width = idx===0? 15 : 12;
  });

  // Casting / Machining group headers
  ws.mergeCells(`D2:${colLetter(COL.CAST_END)}2`);
  ws.getCell('D2').value = 'Casting Defect';
  ws.getCell('D2').font = { bold:true };
  ws.getCell('D2').alignment = { horizontal:'center', vertical:'middle' };

  ws.mergeCells(`${colLetter(COL.MACH_START)}2:${colLetter(COL.MACH_END)}2`);
  ws.getCell(colLetter(COL.MACH_START)+'2').value = 'Machining Defect';
  ws.getCell(colLetter(COL.MACH_START)+'2').font = { bold:true };
  ws.getCell(colLetter(COL.MACH_START)+'2').alignment = { horizontal:'center', vertical:'middle' };

  // Total Reprocess
  ws.mergeCells(`${colLetter(COL.TOTAL_REPROCESS)}2:${colLetter(COL.TOTAL_REPROCESS)}3`);
  ws.getCell(colLetter(COL.TOTAL_REPROCESS)+'2').value = 'TOTAL REPROCESS';
  ws.getCell(colLetter(COL.TOTAL_REPROCESS)+'2').font = { bold:true };
  ws.getCell(colLetter(COL.TOTAL_REPROCESS)+'2').alignment = { horizontal:'center', vertical:'middle' };
  ws.getColumn(COL.TOTAL_REPROCESS).width = 14;

  // OK column
  ws.mergeCells(`${colLetter(COL.OK)}2:${colLetter(COL.OK)}3`);
  ws.getCell(colLetter(COL.OK)+'2').value = 'OK';
  ws.getCell(colLetter(COL.OK)+'2').font = { bold:true };
  ws.getCell(colLetter(COL.OK)+'2').alignment = { horizontal:'center', vertical:'middle' };
  ws.getCell(colLetter(COL.OK)+'2').fill = {
    type:'pattern', pattern:'solid', fgColor:{argb:'FFFF9900'}
  };
  ws.getColumn(COL.OK).width = 12;

  // % columns
  [
    [COL.RW_PCT,'Rework %'],
    [COL.CP_PCT,'Casting Reprocess %'],
    [COL.MP_PCT,'Machining Reprocess %']
  ].forEach(([c,label])=>{
    ws.mergeCells(`${colLetter(c)}2:${colLetter(c)}3`);
    ws.getCell(colLetter(c)+'2').value = label;
    ws.getCell(colLetter(c)+'2').font = { bold:true, size:10 };
    ws.getCell(colLetter(c)+'2').alignment = { wrapText:true, horizontal:'center', vertical:'middle' };
    ws.getCell(colLetter(c)+'2').fill = {
      type:'pattern', pattern:'solid', fgColor:{argb:'FFFFFF00'}
    };
    ws.getColumn(c).width = 12;
  });

  // Sub-headers row 3: castingList in green, machiningList in yellow
  castingList.forEach((item,i)=>{
    const cell = ws.getCell(3, COL.CAST_START + i);
    cell.value = `${item.code} - ${item.label}`;
    cell.alignment = { textRotation: 90, wrapText:true, horizontal:'center', vertical:'bottom' };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FF00B050'} };
    cell.font = { color:{argb:'FFFFFFFF'} };
    ws.getColumn(COL.CAST_START + i).width = 4;
  });
  machiningList.forEach((item,i)=>{
    const cell = ws.getCell(3, COL.MACH_START + i);
    cell.value = `${item.code} - ${item.label}`;
    cell.alignment = { textRotation: 90, wrapText:true, horizontal:'center', vertical:'bottom' };
    cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFFFFF00'} };
    cell.font = { color:{argb:'FF000000'} };
    ws.getColumn(COL.MACH_START + i).width = 4;
  });

  // 3) Data rows
  rows.forEach((r, idx)=>{
    const R = 4 + idx;
    ws.getCell(R, COL.MODEL).value    = r.model;
    ws.getCell(R, COL.CUSTOMER).value = r.customer;
    ws.getCell(R, COL.TOTAL).value    = r.totalChecked;

    // fill casting
    castingList.forEach((item,i)=>{
      ws.getCell(R, COL.CAST_START + i).value = r[item.label];
    });
    // machining
    machiningList.forEach((item,i)=>{
      ws.getCell(R, COL.MACH_START + i).value = r[item.label];
    });

    // TOTAL REPROCESS
    ws.getCell(R, COL.TOTAL_REPROCESS).value = r.rework;
    // OK
    ws.getCell(R, COL.OK).value = r.ok;
    // calculate sums
    const totalCast = castingList.reduce((sum,sub) => sum + (r[sub]||0), 0);
    const totalMach = machiningList.reduce((sum,sub) => sum + (r[sub]||0), 0);
    const okVal      = r.ok || 1;

    // percentages (rounded whole)
    ws.getCell(R, COL.RW_PCT).value = Math.round(r.rework / okVal * 100) + '%';
    ws.getCell(R, COL.CP_PCT).value = Math.round(totalCast / okVal * 100) + '%';
    ws.getCell(R, COL.MP_PCT).value = Math.round(totalMach / okVal * 100) + '%';
  });

  // freeze panes
  ws.views = [{ xSplit: 2, ySplit: 3 }];

  // --- C) write to buffer + email ---
  const buffer = await wb.xlsx.writeBuffer();

    // Step 2: Nodemailer setup
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
        user: 'shashankphogat26@gmail.com',
        pass: 'ednf cczi tebu rtfu', // Use App Password for Gmail
    },
  });

    // Step 3: Send email with Excel attachment
    const mailOptions = {
        from: 'shashankphogat26@gmail.com',
        to: 'nisar@nsidentics.com',
        subject: 'Incoming Inspection Report',
        text: 'Please find the attached inspection report.',
        attachments: [
          {
            filename: 'Incoming_Inspection_Report.xlsx',
            content:  buffer,
            contentType:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          },
        ],
      };
  
      await transporter.sendMail(mailOptions);

        res.status(200).json({ message: 'Inspection report generated and sent successfully!' });
    }
    catch(error) {
        console.error(error);
        res.status(500).json({ message: 'Error generating inspection report' });
      }
})

router.get('/api/generate-pareto-report', async (req, res) => {
  try {
    // — A) Aggregate casting and machining rework counts —
    const [ castingAgg, machiningAgg ] = await Promise.all([
      partsModel.aggregate([
        { $match: { defectType: 'casting'   } },
        { $group: { _id: '$defectSubType', qty: { $sum: '$rework' } } },
        { $sort: { qty: -1 } }
      ]),
      partsModel.aggregate([
        { $match: { defectType: 'machining' } },
        { $group: { _id: '$defectSubType', qty: { $sum: '$rework' } } },
        { $sort: { qty: -1 } }
      ])
    ]);

    // — B) Build Pareto data arrays with cumulative sums & percents —
    function makePareto(arr) {
      let cum = 0;
      const total = arr.reduce((s,r) => s + r.qty, 0) || 1;
      return arr.map((r,i) => {
        cum += r.qty;
        return {
          sr:     i + 1,
          name:   r._id,
          qty:    r.qty,
          cumQty: cum,
          cumPct: (cum / total * 100).toFixed(2) + '%'
        };
      });
    }
    const castingData   = makePareto(castingAgg);
    const machiningData = makePareto(machiningAgg);

    // — C) Helper to build one workbook for a dataset —
    async function buildWorkbook(data, title) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet(title);

      // Table headers
      const headers = ['Sr No', 'Defect Name', 'Defect Qty', 'Cumm. Qty', 'Cumm. %'];
      headers.forEach((h, i) => {
        const cell = ws.getCell(1, i+1);
        cell.value = h;
        cell.font = { bold: true };
        ws.getColumn(i+1).width = [8, 30, 12, 12, 12][i];
      });

      // Data rows
      data.forEach((r, idx) => {
        const row = ws.getRow(2 + idx);
        row.getCell(1).value = r.sr;
        row.getCell(2).value = r.name;
        row.getCell(3).value = r.qty;
        row.getCell(4).value = r.cumQty;
        row.getCell(5).value = r.cumPct;
      });

      // Pareto chart via QuickChart
      const qc = new QuickChart()
      .setConfig({
        type: 'bar',
        data: {
          labels:    data.map(r => r.name),
          datasets: [
            {
              type: 'bar',
              label: 'Qty',
              data:  data.map(r => r.qty),
              backgroundColor: 'rgba(0, 112, 192, 0.7)'
            },
            {
              type: 'line',
              label: 'Cumulative %',
              yAxisID: 'pct',
              data:  data.map(r => parseFloat(r.cumPct)),
              borderColor: 'orange',
              backgroundColor: 'transparent',
              fill: false,
              pointBackgroundColor: 'orange',
              tension: 0.3
            }
          ]
        },
        options: {
          scales: {
            y:   { beginAtZero: true },
            pct: {
              type: 'linear', position: 'right',
              beginAtZero: true, max: 100,
              ticks: { callback: v => v + '%' }
            }
          },
          plugins: { legend: { position: 'top' } },
          title:   { display: true, text: title }
        }
      })
      .setWidth(600)
      .setHeight(300)
      .setBackgroundColor('white')
      .setVersion('3.7.1');    // ensure v3 syntax compatibility

      const imgBuffer = await qc.toBinary();

      ws.mergeCells(0, 6, 0, 11);  
      ws.getCell(1, 7).value = title;  // cell G1
      ws.getCell(1, 7).font  = { bold:true, size:14 };
      ws.getCell(1, 7).alignment = { horizontal:'center' };


      const imgId = wb.addImage({ buffer: imgBuffer, extension: 'png' });
      // place chart at col G,row1
      ws.addImage(imgId, { tl: { col: 6, row: 0 }, ext: { width: 600, height: 300 } });

      return wb.xlsx.writeBuffer();
    }

    // — D) Build both workbooks in parallel —
    const [ bufCast, bufMach ] = await Promise.all([
      buildWorkbook(castingData,   'Casting Defects Pareto'),
      buildWorkbook(machiningData, 'Machining Defects Pareto')
    ]);

    // — E) Send both as email attachments —
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
      user: 'shashankphogat26@gmail.com',
      pass: 'ednf cczi tebu rtfu', // Use App Password for Gmail
      }
    });

    await transporter.sendMail({
      from:    'shashankphogat26@gmail.com',
      to:       'nisar@nsidentics.com',
      subject:  'Daily Pareto Defect Reports',
      text:     'Attached are today’s Pareto defect reports for casting and machining.',
      attachments: [
        {
          filename: 'Pareto_Casting_Defects.xlsx',
          content:  bufCast,
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        },
        {
          filename: 'Pareto_Machining_Defects.xlsx',
          content:  bufMach,
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        }
      ]
    });

    res.json({ success: true, message: 'Reports generated & emailed.' });
  }
  catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});

router.get('/api/send-model-trend', async (req, res) => {
  try {
    // A) Aggregate by model
    const agg = await partsModel.aggregate([
      {
        $group: {
          _id: '$model',
          totalChecked: { $sum: { $add: ['$ok', '$rework', '$reject'] } },
          ok:           { $sum: '$ok' },
          rework:       { $sum: '$rework' }
        }
      },
      { $sort: { totalChecked: -1 } }
    ]);

    // B) Prepare rows
    const rows = agg.map((d, i) => ({
      sr:           i + 1,
      model:        d._id,
      totalChecked: d.totalChecked,
      ok:           d.ok,
      rework:       d.rework
    }));

    // C) Build Excel workbook
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Model Trend');

    // 1) Table headers
    const headers = [
      'Sr No', 'Model',
      'Sum of Total Checked',
      'Sum of OK Qty',
      'Sum of Rework Qty'
    ];
    headers.forEach((h, i) => {
      const cell = ws.getCell(1, i + 1);
      cell.value = h;
      cell.font = { bold: true };
      ws.getColumn(i + 1).width = [8, 20, 18, 14, 16][i];
    });

    // 2) Table data
    rows.forEach(r => {
      const row = ws.getRow(2 + r.sr - 1);
      row.getCell(1).value = r.sr;
      row.getCell(2).value = r.model;
      row.getCell(3).value = r.totalChecked;
      row.getCell(4).value = r.ok;
      row.getCell(5).value = r.rework;
    });

    // 3) Freeze header row
    ws.views = [{ state: 'frozen', ySplit: 1 }];

    // D) Generate grouped‑bar chart via QuickChart
    const qc = new QuickChart()
      .setConfig({
        type: 'bar',
        data: {
          labels: rows.map(r => r.model),
          datasets: [
            {
              label: 'Total Checked',
              data:  rows.map(r => r.totalChecked),
              backgroundColor: 'rgba(192,192,192,0.7)'  // grey
            },
            {
              label: 'OK Qty',
              data:  rows.map(r => r.ok),
              backgroundColor: 'rgba(0,112,192,0.7)'    // blue
            },
            {
              label: 'Rework Qty',
              data:  rows.map(r => r.rework),
              backgroundColor: 'orange'                 // orange
            }
          ]
        },
        options: {
          plugins: {
            legend: { position: 'top' },
            title:  { display: true, text: 'Model Trend' }
          },
          scales: {
            y: { beginAtZero: true }
          }
        }
      })
      .setWidth(900)
      .setHeight(400)
      .setBackgroundColor('white')
      .setVersion('3.7.1');

    const chartBuffer = await qc.toBinary();

    // E) Embed chart in Excel (col G=7 zero-indexed: col 6, row 1)
    const imgId = wb.addImage({ buffer: chartBuffer, extension: 'png' });
    ws.addImage(imgId, {
      tl:  { col: 6, row: 0 },
      ext: { width: 900, height: 400 }
    });

    // F) Email the workbook
    const buffer = await wb.xlsx.writeBuffer();
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
      user: 'shashankphogat26@gmail.com',
      pass: 'ednf cczi tebu rtfu', // Use App Password for Gmail
      }
    });

    await transporter.sendMail({
      from:    'shashankphogat26@gmail.com',
      to:       'nisar@nsidentics.com',
      subject:    'Daily Model Trend Report',
      text:       'Attached is today’s model trend report.',
      attachments: [{
        filename:    'Model_Trend_Report.xlsx',
        content:     buffer,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }]
    });

    res.json({ success: true, message: 'Model trend report sent.' });
  }
  catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});

router.post('/api/hourly-report', async (req, res) => {
  try {
    const { model } = req.body;
    if (!model) {
      return res.status(400).json({ error: 'model is required in request body' });
    }

    // 3) Determine the target date (midnight start → next‑day midnight)
    // const target = date
    //   ? new Date(date)
    //   : new Date();
    const target = new Date();
    const start = new Date(target.getFullYear(), target.getMonth(), target.getDate(), 0, 0, 0);
    const end   = new Date(start);
    end.setDate(end.getDate() + 1);

    // 4) Aggregate by hour
    const hourly = await partsModel.aggregate([
      { $match: { model, createdAt: { $gte: start, $lt: end } } },
      {
        $project: {
          hour:   { $hour: '$createdAt' },
          ok:     1,
          rework: 1
        }
      },
      {
        $group: {
          _id: '$hour',
          okQty:     { $sum: '$ok' },
          reworkQty: { $sum: '$rework' }
        }
      },
      { $sort: { _id: 1 } }
    ]);
    console.log(hourly)

    // 5) Build rows with cumulative OK
    let cumOk = 0;
    const rows = hourly.map((h, idx) => {
      cumOk += h.okQty;
      const hour = h._id;
      const timeLabel = `${hour}:00 to ${hour + 1}:00`;
      const dateLabel = start.toLocaleDateString('en-GB'); // dd/MM/yyyy
      return {
        sr:          idx + 1,
        model,
        okQty:       h.okQty,
        reworkQty:   h.reworkQty,
        cumOkQty:    cumOk,
        timeLabel,
        dateLabel
      };
    });

    // 6) Build Excel workbook
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Hourly Inspection');

    // 6a) Merged title row
    ws.mergeCells('A1:G1');
    ws.getCell('A1').value = 'Hourly inspection report';
    ws.getCell('A1').font = { bold: true, size: 14 };
    ws.getCell('A1').alignment = { horizontal: 'center' };

    // 6b) Table headers in row 2
    const headers = [
      'Sr. No.', 'Model', 'Ok Qty.', 'Rework Qty.', 'Cumm. Ok Qty.', 'Time', 'Date'
    ];
    headers.forEach((h, i) => {
      const cell = ws.getCell(2, i + 1);
      cell.value = h;
      cell.font = { bold: true };
      ws.getColumn(i + 1).width = [8, 15, 10, 12, 14, 20, 14][i];
    });

    // 6c) Data rows start at row 3
    rows.forEach(r => {
      const R = 2 + r.sr;
      ws.getCell(R, 1).value = r.sr;
      ws.getCell(R, 2).value = r.model;
      ws.getCell(R, 3).value = r.okQty;
      ws.getCell(R, 4).value = r.reworkQty;
      ws.getCell(R, 5).value = r.cumOkQty;
      ws.getCell(R, 6).value = r.timeLabel;
      ws.getCell(R, 7).value = r.dateLabel;
    });

    // 6d) Freeze the top 2 rows
    ws.views = [{ state: 'frozen', ySplit: 2 }];

    // 7) Email via Nodemailer
    const buffer = await wb.xlsx.writeBuffer();
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
      user: 'shashankphogat26@gmail.com',
      pass: 'ednf cczi tebu rtfu', // Use App Password for Gmail
      }
    });

    await transporter.sendMail({
      from:    '"Hourly Bot" <shashankphogat26@gmail.com>',
      to:       'nisar@nsidentics.com',
      subject:  `Hourly Inspection Report for ${model} (${start.toLocaleDateString()})`,
      text:     `Please find attached the hourly inspection report for ${model} on ${start.toLocaleDateString()}.`,
      attachments: [{
        filename: `Hourly_Report_${model}_${String(target) || start.toISOString().slice(0,10)}.xlsx`,
        content:  buffer,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }]
    });

    res.json({ success: true, message: 'Hourly report generated & emailed.' });
  }
  catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: err.message });
  }
});


module.exports=router