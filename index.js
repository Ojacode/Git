const express=require('express')
const multer= require("multer")
const reader= require('xlsx')
const ExcelJS=require('exceljs');

const app=express()
const port=process.env.PORT || 3000


app.use(express.json())

app.use('/publicfiles',express.static(__dirname + '/publicfiles'))


app.get('/readexcelfile',(req,res)=>{
  let filename=req.query.filename;
  let data=[]
  try {
    const file=reader.readFile('publicfiles/' + filename + ".xlsx")
    const sheetNames=file.SheetNames

    for(let i=0;i<sheetNames.length;i++)
    {
      const arr= reader.utils.sheet_to_json(file.Sheets[sheetNames[i]])
      arr.forEach((res)=>{
        data.push(res)
      })
    }
    res.send(data);

  console.log(data)
  let i = 1;
  let sum1=0;
  let sum2=0;
  let sum3=0;
  let result={}
  for(i=0;i<256;i++)
  {
    
    if(data[i].CO1>=7)
    {
      sum1=sum1+1;
    }
    if(data[i].CO1>=6)
    {
      sum2=sum2+1;
    }
    if(data[i].CO1>=4)
    {
      sum3=sum3+1;
    }
    // console.log(data[i].CO1);

  }
  } catch (e) {
    res.send(e)
  }
})

var storage = multer.diskStorage({
    destination: 'publicfiles',
    filename: function (req, file, callback) {
        callback(null, file.originalname);
    }
});


const upload=multer({
  dest:'publicfiles',
  storage:storage,
  limits :{
    fileSize:1000000
  },
  fileFilter(req,file,cb){

    if(!file.originalname.match('xlsx')){
      return cb(new Error('Please upload an image'))
    }
    
    cb(undefined,true)
  }
})

const errorMiddleware =(req,res,next) =>{
    throw new Error('From my middleware')
} 

app.post('/upload',upload.single('upload'), async(req,res)=>{

  res.send()
},(error,req,res,next)=>{
   res.status(400).send({error:error.message})
})


app.post('/sheet',async(req,res)=>{
  try{
   // Requiring module
// const reader = require('xlsx')

// Reading our test file
let workbook=new ExcelJS.Workbook()
await workbook.xlsx.readFile('./test.xlsx')
let worksheet=workbook.getWorksheet("Sheet1")

let distinction = worksheet.getRow(3);
let firstclass= worksheet.getRow(4);
let secondclass = worksheet.getRow(5);

distinction.getCell(2).value = 2;
distinction.getCell(3).value = 2;
distinction.getCell(5).value = 2;
distinction.getCell(6).value = 2;

firstclass.getCell(2).value = 2;
firstclass.getCell(3).value = 2;
firstclass.getCell(5).value = 2;
firstclass.getCell(6).value = 2;


secondclass.getCell(2).value = 2;
secondclass.getCell(3).value = 2;
secondclass.getCell(5).value = 2;
secondclass.getCell(6).value = 2;


distinction.commit()
firstclass.commit()
secondclass.commit()

 await workbook.xlsx.writeFile('./test.xlsx');

    res.send('done');
  }
  catch(e){
    res.status(400).send(e);
  }
});


app.listen(port,()=>{
  console.log("Server is up on port "+ port)
})

