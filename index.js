const express    = require('express');
const mongoose   = require('mongoose');
const bodyParser = require('body-parser');
const path       = require('path');
const XLSX       = require('xlsx');
const excelJS = require('exceljs');
const multer     = require('multer');


//multer
var storage = multer.diskStorage({
    destination: function (req, file, cb) {
      cb(null, './public/uploads')
    },
    filename: function (req, file, cb) {
      cb(null, file.originalname)
    }
  });
  
  var upload = multer({ storage: storage });

//connect to database
mongoose.connect('mongodb+srv://NewOne:Durgesh%402022@cluster0.lvsxb5w.mongodb.net/inotebooktest',
{
    useNewUrlParser:true,
    useUnifiedTopology: true
})
.then(()=>{console.log('connected to database successfully')})
.catch((error)=>{console.log('error',error)});

//init app
var app = express();

//set the template engine
app.set('view engine','ejs');

//fetch data from the request
app.use(bodyParser.urlencoded({extended:false}));

//static folder path
app.use(express.static(path.resolve(__dirname,'public')));

//collection schema
var fileSchema = new mongoose.Schema({
    // ORDERNUMBER:Number,
    // QUANTITYORDERED:Number,
    // PRICEEACH:Number,
    SALES:Number,
    // ODERDATE:Date,
    STATUS:String,
    PRODUCTLINE:String,
    TERRITORY:String,
});

var Filedata = mongoose.model('fileData',fileSchema);

// endpoints

app.get('/',(req,res)=>{
   Filedata.find((err,data)=>{
       if(err){
           console.log(err)
       }else{
           if(data!=''){
               res.render('home',{result:data});
           }else{
               res.render('home',{result:{}});
           }
       }
   });
});

app.post('/',upload.single('excel'),(req,res)=>{
  var workbook =  XLSX.readFile(req.file.path);
  var sheet_namelist = workbook.SheetNames;
  var x=0;
  sheet_namelist.forEach(element => {
      var xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_namelist[x]]);
      Filedata.insertMany(xlData,(err,data)=>{
          if(err){
              console.log(err);
          }else{
              console.log(data);
          }
      })
      x++;
  });
  res.redirect('/');
});

app.get('/export-users', async (req, res)=>{
     try {
         
        const workbook = new excelJS.Workbook();
        const worksheet =  workbook.addWorksheet("data");
   
        worksheet.columns =[
            // {header:'S.no', key:"s_no"},
            {header:'PRODUCTLINE', key:"PRODUCTLINE"},
            {header:'TERRITORY', key:"TERRITORY"},
            {header:'STATUS', key:"STATUS"},
            {header:'NetShippedUnits',  key:"STATUS"},
            {header:'NetSale', key:"SALES"},
        ];

        // let count = 1; // adding serial number

        const userdata = await Filedata.find({});
 
        userdata.forEach((user)=>{
            // user.s_no = count;
            // const {STATUS, } = user
            worksheet.addRow(user);z
            // count ++;
        })

        worksheet.getRow(1).eachCell((cell)=>{
            cell.font = {bold:true};
        });

        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheatml.sheet"
        );

        res.setHeader("Content-Disposition", `attachment; filename = users.xlsx`);

        return workbook.xlsx.write(res).then(()=>{
            res.status(200);
        })

     } catch (error) {
        console.log(error.message)
     }
})

//assign port
const port = 3000;
app.listen(port,()=>console.log(`server connected to http://localhost:${port}`));
