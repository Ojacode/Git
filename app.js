const express=require('express')
const multer= require("multer")
const reader= require('xlsx')

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

app.listen(port,()=>{
  console.log("Server is up on port "+ port)
})
