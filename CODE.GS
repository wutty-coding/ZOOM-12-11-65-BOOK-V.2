var folderPDF ='xxxx'// ที่เก็บไฟล์ PDF
var idslides ='xxxx'// แม่แบบ
var idToken ='xxxx' // โทเค่น แจ้งเตือนเข้าไลน์


function doGet(){
  var template = HtmlService.createTemplateFromFile('index');
  template.data = getSheetData();
  var output = template.evaluate()
  .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  return output;

}

function getSheetData(){
 return ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('menu').getDataRange().getDisplayValues().slice(1);

}

function include(filename){
   return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

/** @Get URL */
function getURL() {
  return ScriptApp.getService().getUrl();
}

/** @ค้นหาข้อมูลจากเบอร์โทร */
function findRecord(record) {
  var ss = SpreadsheetApp.getActive().getSheetByName('user')
  var data = ss.getDataRange().getValues()
  var id = data.map(r => r[0]) // ดึงค่ามารอเพื่อตรวจสอบ
  var index = id.indexOf(record)
  var output = {}
  if (index >= 1) {
    data = ss.getRange(index + 1, 1, 1, ss.getLastColumn()).getValues()[0]
    output.result = true
    output.data = data
  } else {
    output.result = false
  }
  return output
}


/** @ค้นหาแยกห้องเรียน */
function getData2(idserach){
  var data = SpreadsheetApp.getActive().getSheetByName("data").getDataRange().getDisplayValues().slice(1)
  var output = []
  data.forEach(function(f){
    if(~f.indexOf(idserach)){
output.push(f)
    }
  })
  return output

}


function deleteId2(id){
   let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('data')
  let datasheet = sheet.getDataRange().getDisplayValues()
  let idcheck = datasheet.map(r => r[0])
  var index = idcheck.indexOf(id)
  Logger.log(index)
  if(index != -1){
     sheet.deleteRow(index+1); 
    return true;
  }
}


function saveData(data) {
  var folder = DriveApp.getFolderById("1FMrU-xwv1wZ9WGvB_PcR1_1keYHnEVOn")// เก็บรูป 6 รูป

  /**วันรายงาน เครดิตครูนิดเดียว*/
  var date = data.input1.split("-")
  var dateThai = Number(date[2]); //วัน
  var mounthText = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]; //อาเรย์ชุดเดือน 
  var mounthThai = mounthText[Number(date[1])]; // หาค่าเดือนเป็นไทย
  var yearThai = Number(date[0]) + 543 // ปี
  var dayThai = dateThai + ' ' + mounthThai + ' พ.ศ. ' + yearThai

  if (data.imgdata.length > 0) {
    const fileimg1 = data.imgdata.split(",")[1];
    const img1 = Utilities.base64Decode(fileimg1);
    const idPNG1 = Utilities.newBlob(img1).setName(data.inputname + ".png");
    var pic1 = folder.createFile(idPNG1).getId();
    var image_url1 = "https://lh3.googleusercontent.com/d/" + pic1
  } else {
    image_url1 = ""
  }



  if (data.imgdata2.length > 0) {
    const fileimg2 = data.imgdata2.split(",")[1];
    const img2 = Utilities.base64Decode(fileimg2);
    const idPNG2 = Utilities.newBlob(img2).setName(data.inputname + ".png");
    var pic2 = folder.createFile(idPNG2).getId();
    var image_url2 = "https://lh3.googleusercontent.com/d/" + pic2
  } else {
    image_url2 = ""
  }


  if (data.imgdata3.length > 0) {
    const fileimg3 = data.imgdata3.split(",")[1];
    const img3 = Utilities.base64Decode(fileimg3);
    const idPNG3 = Utilities.newBlob(img3).setName(data.inputname + ".png");
    var pic3 = folder.createFile(idPNG3).getId();
    var image_url3 = "https://lh3.googleusercontent.com/d/" + pic3
  } else {
    image_url3 = ""
  }

  if (data.imgdata4.length > 0) {
    const fileimg4 = data.imgdata4.split(",")[1];
    const img4 = Utilities.base64Decode(fileimg4);
    const idPNG4 = Utilities.newBlob(img4).setName(data.inputname + ".png");
    var pic4 = folder.createFile(idPNG4).getId();
    var image_url4 = "https://lh3.googleusercontent.com/d/" + pic4
  } else {
    image_url4 = ""
  }

  if (data.imgdata5.length > 0) {
    const fileimg5 = data.imgdata5.split(",")[1];
    const img5 = Utilities.base64Decode(fileimg5);
    const idPNG5 = Utilities.newBlob(img5).setName(data.inputname + ".png");
    var pic5 = folder.createFile(idPNG5).getId();
    var image_url5 = "https://lh3.googleusercontent.com/d/" + pic5
  } else {
    image_url5 = ""
  }

  if (data.imgdata6.length > 0) {
    const fileimg6 = data.imgdata6.split(",")[1];
    const img6 = Utilities.base64Decode(fileimg6);
    const idPNG6 = Utilities.newBlob(img6).setName(data.inputname + ".png");
    var pic6 = folder.createFile(idPNG6).getId();
    var image_url6 = "https://lh3.googleusercontent.com/d/" + pic6
  } else {
    image_url6 = ""
  }

  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data')
  ss.appendRow([
    "",
    dayThai,
    "'"+data.searchTel,
    data.idplan,
    data.imageuser,
    data.inputname,
    data.input2,
    data.input3,
    image_url1,
    image_url2,
    image_url3,
    image_url4,
    image_url5,
    image_url6,

  ])
  runPDF()
}




function runPDF(){

  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data')
  var pdfFolder = DriveApp.getFolderById(folderPDF) // ID โฟลเดอร์ JPG เก็บไฟล์
 
  let copyFile = DriveApp.getFileById(idslides).makeCopy(), // คัดลอกไฟล์สำเนาจากสไลด์
      copyID = copyFile.getId(), // คัดลอกไอดี
      copyDoc = SlidesApp.openById(copyID) // ไฟล์ก๊อปปี้
 
  let headerRow = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues(), // ส่วนเฮดเดอร์ของแถว
    item = ss.getRange(ss.getLastRow(), 1, 1, ss.getLastColumn()).getDisplayValues(), // ข้อมูลแถวล่าสุด
    columnIndex = 0 // Index คอลัมภ์เริ่มต้น
 
  /** ส่วนแทนที่ข้อความทั้งหมด */
  for (; columnIndex < headerRow[0].length; columnIndex++) {
      copyDoc.replaceAllText('{' + headerRow[0][columnIndex] + '}', item[0][columnIndex]) // แทนที่ส่วนที่เป็นข้อความทั้งหมด
  }
 
 /** ส่วนแทนที่รูปภาพตามชื่อหัวคอลัมภ์ รูปแบบที่ 3 (กำหนดรูปโดยให้วนลูปหาทุกรูปจากทุกสไลด์) */
  //============ ชุดวิ่งหาและแทนที่ภาพ เครดิต ครูสมพงษ์ โพคาศรี ========
  var imageData = ss.getRange(ss.getLastRow(), 1, 1, ss.getLastColumn()).getDisplayValues()[0]
  var data = ss.getDataRange().getDisplayValues()
  var headers = [...data[0]]
 
  let images = []
  copyDoc.getSlides().forEach(slide => images.push(...slide.getImages()))
  images.forEach(images => {
    let title = images.getTitle()
    for (var img = 0; img < headers.length; img++) {
      if (headers[img] == title) {
        var imageSplit = imageData[img].split("d/")
        var driveImage = Drive.Files.get(imageSplit[1]).thumbnailLink.replace('220', '2000')
        images.replace(driveImage, true); //ขนาดเท่าเดิมและขนาดกว้างหรือยาวด้านใดด้านนึงเป็นหลัก 
      }
    }
  })

  copyDoc.saveAndClose() // บันทึกและปิดไฟล์ก๊อปปี้
 
  // //===========ส่วนการสร้างไฟล์ PDF =========================
  // var newFile = pdfFolder.createFile(copyFile.getAs(MimeType.PDF)) // สร้างไฟล์ใหม่เป็น PDF
  // var pdfView = newFile.setName(item[0][0] + item[0][1]).getUrl() // สร้างลิงค์ PDF แบบวิว


   //===========ส่วนการสร้างไฟล์ JPG =========================
  var jpgOut1 = "" //กำหนดาค่าว่างของลิงก์ไฟล์ภาพ jpg
  var newJPGFile = "" //กำหนดาค่าว่างของชื่อไฟล์
  const srcSlides = copyDoc.getSlides();
  srcSlides.forEach((s, i) => {
    const url = Slides.Presentations.Pages.getThumbnail(copyID, s.getObjectId(), { "thumbnailProperties.mimeType": "PNG" }).contentUrl;
    const blob = UrlFetchApp.fetch(url).getAs(MimeType.JPEG);
    newJPGFile = pdfFolder.createFile(blob.setName(item[0][0] + item[0][1] ));
    jpgOut1 = newJPGFile.getId() //ไอดีของไฟล์ภาพ jpg
    ucjpg = "https://lh3.googleusercontent.com/d/" + jpgOut1
  });


  // ส่วนการแทรกข้อมูลลงในชีต
  ss.getRange(ss.getLastRow(),ss.getLastColumn()).setValue(ucjpg) // เซ็ตลิงค์ไฟล์ PDF ลงในคอลัมภ์ที่...
  copyFile.setTrashed(true) // ลบไฟล์สำเนาลงถังขยะ


  /**ชุดส่งไลน์เป็นลิ้งค์และภาพ */
      var idnewJPG=newJPGFile.getId() //ไอดีของไฟล์ภาพ jpg
      var imgLine = DriveApp.getFileById(idnewJPG).getBlob()
      let msg="แจ้งรายงานสร้างวารสาร"
      let token = idToken// ไอดี Token แจ้งไลน์
      sendNotify(msg,token,imgLine)


}


function sendNotify(msg,token,imgLine){
let payloadJson = {
  "message":msg,    
  "imageFile": imgLine
    };    
    let options = {
        "method": "post",
        "payload": payloadJson,
        "headers": {
            "Authorization": "Bearer " + token
        }
    };
    UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}
 
 



