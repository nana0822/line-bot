var CHANNEL_ACCESS_TOKEN = 'ytFDV+ldre+k6/kMgHSqUhmGOHSS270sQVUpOG7x2foeJz9oCOLkMOlGLIyMsTM8VpdtCNdDE24Oqunbp6n7s34+qXkEuWj2NQ/RxkwQV9IRumYB7MlaO4wkLS+WyH2Q95Augr2haP1/Qi3ys1itYwdB04t89/1O/w1cDnyilFU='; 
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('応答文');
var log_sheet = ss.getSheetByName('ログ');

var line_endpoint_reply = 'https://api.line.me/v2/bot/message/reply';
var line_endpoint_push = 'https://api.line.me/v2/bot/message/push';

var reception_message = 1;
var type = 2;
var template_type = 3;
var choices = 4;
var text_title = 5;
var sub_text = 6;
var img_url = 7;
var btn = {"btn1":{"type":8, "label":9, "text":10}, "btn2":{"type":11, "label":12, "text":13}, "btn3":{"type":14, "label":15, "text":16}, btn4:{"type":17, "label":18, "text":19}};
var continuous = 20;

var log = [];

function doPost(e) {
  try{
    log.push(Utilities.formatDate( new Date(), 'Asia/Tokyo', 'yyyy年M月d日'))
    var json = JSON.parse(e.postData.contents);
    log.push(JSON.stringify(json));
    var reply_token= json.events[0].replyToken;
    log.push("token get = " + reply_token);
    if (typeof reply_token === 'undefined') {
      return;
    }
    var message = json.events[0].message.text;
    log.push("message get = "+ message);
    var messages = new Array();
    var type = "";
    var row = findRow(message);
    if(row === 0){
      row = 2;
    }
    log.push("row get = " + row);
    var continu_row = row;
    do{
      type = sheet.getRange(row,template_type).getValue();
      log.push("type get = " + type);
      if(type === 'text'){
        messages.push(text_message(row)); 
      }else if(type === 'sticker'){
        messages.push(sticker_message(row));
      }else if(type === 'carousel'){
        messages.push(carousel_message(row));
      }else if(type === 'buttons'){
        messages.push(buttons_message(row));
      }else if(type === 'location'){
        messages.push(location_message(row));
      }else if(type === 'image'){
        messages.push(image_message(row));
      }else if(type === 'video'){
        messages.push(video_message(row));
      }else if(type === 'confirm'){
        messages.push(confirm_message(row));
      }else if(type === 'quickReply'){
        messages.push(quickReply_message(row));
      }
      var continu = sheet.getRange(continu_row, continuous++).getValue();
      message = continu;
      if(continu){
        row = findRow(message);
        if(row === 0){
          row = 2; 
        }
      }
    }while(message);
    log.push(JSON.stringify(messages)); 
    
    UrlFetchApp.fetch(line_endpoint_reply, {
      'headers': {
        'Content-Type': 'application/json; charset=UTF-8',
        'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
      },
      'method': 'post',
      'payload': JSON.stringify({
        "replyToken" : reply_token, 
        "messages" : messages
      })
    });
  }catch(e){
    log.push(type + "にてエラーが発生しました。 " + e);
    log_write();
  }
  return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
}

function messages_get(row){
  return sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function text_message(row){
  var val = sheet.getRange(row,text_title).getValue();
  return {
    "type" : "text",
    "text" : val,
    "wrap": true
  };
}

function sticker_message(row){
  var messages = messages_get(row);
  var packageId = messages[text_title-1];
  var stickerId = messages[sub_text-1];
  return {
    "type": "sticker",
    "packageId": packageId,
    "stickerId": stickerId
  };
}

function carousel_message(row){
  var messages = messages_get(row);
  var num = Number(messages[choices-1]);
  var columns = [];
  var title;
  var subtitle;
  do{
    title = messages[text_title-1];
    subtitle = messages[sub_text-1];
    var url = messges[img_url-1];
    var actions = [];
    for(var i=1;i<=num;i++){
      var button_name = "btn" + i;
      var b_type = messages[btn[button_name]["type"]-1];
      var label = messages[btn[button_name]["label"]-1];
      var text = messages[btn[button_name]["text"]-1];
      if(b_type === "uri"){
        var val = {type:b_type, 
                   label:label, 
                   uri:text,
                  };
      }else{
        var val = {type:b_type, 
                   label:label, 
                   text:text,
                  };
      }
      actions.push(val);
    }
    var column = [];
    var column = {
      "thumbnailImageUrl": url,
      "title": title,
      "text": subtitle,
      "actions": actions
    };
    if(!title){
      delete column["title"];
    }
    columns.push(column);
    row++;
  }while(sheet.getRange(row,reception_message).getValue() === sheet.getRange(row-1,reception_message).getValue());
  return {
    "type": "template",
    "altText": title !== "" ? title: subtitle,
    "template": {
      "type": "carousel",
      "columns": columns
    }
  };
}

function buttons_message(row){
  var messages = messages_get(row);
  var num = Number(messages[choices-1]);
  var title = messages[text_title-1];
  var subtitle = messages[sub_text-1];
  var url = messages[img_url-1];
  var actions = [];
  for(var i=1;i<=num;i++){
    var button_name = "btn" + i;
    var b_type = messages[btn[button_name]["type"]-1];
    var label = messages[btn[button_name]["label"]-1];
    var text = messages[btn[button_name]["text"]-1];
    if(b_type === "uri"){
      var val = {type:b_type, 
                 label:label, 
                 uri:text,
                };
    }else{
      var val = {type:b_type, 
                 label:label, 
                 text:text,
                };
    }
    actions.push(val);
  }
  var re;
  if(url){
    re = { "type": "template", 
          "altText": title !== "" ? title:subtitle, 
          "template": { "type": "buttons",
                       "actions": actions, 
                       "thumbnailImageUrl": url,
                       "title": title, 
                       "text": subtitle,
                      },
         };
  }else{
    re = { "type": "template", 
          "altText": title !== "" ? title:subtitle, 
          "template": { "type": "buttons", 
                       "actions": actions, 
                       "title": title, 
                       "text": subtitle,
                      },
         };
  }
  if(!title){
    delete re.template["title"];
  }
  return re;
}

function location_message(row){
  var messages = messages_get(row);
  var title = messages[text_title-1];
  title = title;
  var address = messages[sub_text-1];
  var positions = messages[img_url-1];
  var position = positions.split("/");
  return { 
    "type": "location", 
    "title": title, 
    "address": address, 
    "latitude": position[0],
    "longitude": position[1] 
  };
}

function image_message(row){
  var img = sheet.getRange(row,img_url).getValue();
  return {
    "type": "image",
    "originalContentUrl": img,
    "previewImageUrl": img,
  };
}

function video_message(row){
  var messages = messages_get(row);
  var video = messages[img_url-1];
  var thum_url = messages[sub_text-1];
  return {
    "type": "video",
    "originalContentUrl": video,
    "previewImageUrl": thum_url
  };
}

function confirm_message(row){
  var messages = messages_get(row);
  var title = messages[text_title-1];
  var actions = [];
  var actions = [];
  for(var i=1;i<=2;i++){
    var button_name = "btn" + i;
    var b_type = messages[btn[button_name]["type"]-1];
    var label = messages[btn[button_name]["label"]-1];
    var text = messages[btn[button_name]["text"]-1];
    if(b_type === "uri"){
      var val = {type:b_type, 
                 label:label, 
                 uri:text,
                };
    }else{
      var val = {type:b_type, 
                 label:label, 
                 text:text,
                };
    }
    actions.push(val);
  }
  if(title){
    return {
      "type": "template",
      "altText": "confirm",
      "template": {
        "type": "confirm",
        "actions": actions,
        "text": title
      }
    };
  }else{
    return {
      "type": "template",
      "altText": "confirm",
      "template": {
        "type": "confirm",
        "actions": actions,
      }
    };
  }
}

// クイックリプライ
function quickReply_message(row){
  var messages = messages_get(row);
  var val = messages[text_title-1];
  var quick_text = messages[img_url-1];
  var quick_texts = quick_text.split(",");
  var items = [];
  for(var i = 0; i < quick_texts.length; i++){
    var item = {
      type: "action",
      action: {
        type: "message",
        label: quick_texts[i],
        text: quick_texts[i]
      }
    };
    items.push(item);
  }
  return {
    "type" : "text",
    "text" : val,
    "quickReply": {
      "items": items
    }
  };
}

function findRow(text){
  var items = sheet.getRange(2, 1, sheet.getLastRow()).getValues();
  for(var i = 0; i < items.length ; i++){
    if(items[i][0] === text){
      return i + 2;
    }
    if(items[i][0] === ""){
      break;
    }
  }
  return 0;
}

function log_write(){
  log_sheet.appendRow(log);
}

