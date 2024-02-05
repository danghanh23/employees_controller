function doGet(e) {
  if (!e.parameter.page) 
  {
    var htmlOutput =  HtmlService.createTemplateFromFile('Index');
    htmlOutput.message = '';
    return htmlOutput.evaluate();
  }
  else if(e.parameter['page'] == 'Link 1')
  {
    Logger.log(JSON.stringify(e));
    var htmlOutput =  HtmlService.createTemplateFromFile('Link 1');
    htmlOutput.office = e.parameter['office'].toString();
    htmlOutput.username = e.parameter['username'].toString();
    return htmlOutput.evaluate();  
  }
  else if(e.parameter['page'] == 'Link 2')
  {
    var htmlOutput =  HtmlService.createTemplateFromFile('Link 2');
    htmlOutput.office = e.parameter['office'].toString();
    htmlOutput.username = e.parameter['username'].toString();
    return htmlOutput.evaluate();  
  } 
  else if(e.parameter['page'] == 'Index')
  {
    var htmlOutput =  HtmlService.createTemplateFromFile('Index');
    htmlOutput.message = e.parameter['message'];
    return htmlOutput.evaluate();  
  }   
}


function getUrl() {
 var url = ScriptApp.getService().getUrl();
 return url;
}

function getUsers(input_office0, input_username0) {
  var input_office =  input_office0.trim();
  var input_username =  input_username0.trim();

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees'); 
  var data = sheet.getDataRange().getValues();

  var users = [];
  if(input_office == "" && input_username == ""){
        for (var i = 1; i < data.length; i++) {
            var user = {
              office: data[i][1],
              user_name: data[i][2],
            };
            users.push(user);
        }

  }else if(input_office == ""){
        for (var i = 1; i < data.length; i++) {
          var office = (data[i][1] || '').toString().toLowerCase(); 
          var user_name = (data[i][2] || '').toString().toLowerCase(); 

          if ((user_name.toLowerCase()).includes(input_username.toLowerCase())) {
                var user = {
                  office: data[i][1],
                  user_name: data[i][2],
                };
                users.push(user);
              }
        }
  }else if(input_username == ""){
          for (var i = 1; i < data.length; i++) {
          var office = (data[i][1] || '').toString().toLowerCase(); 
          var user_name = (data[i][2] || '').toString().toLowerCase(); 

          if ((office.toLowerCase()).includes(input_office.toLowerCase())) {
                var user = {
                  office: data[i][1],
                  user_name: data[i][2],
                };
                users.push(user);
              }
        }
  } else{
          for (var i = 1; i < data.length; i++) {
          var office = (data[i][1] || '').toString().toLowerCase(); 
          var user_name = (data[i][2] || '').toString().toLowerCase(); 

          if ((office.toLowerCase()).includes(input_office.toLowerCase()) &&  (user_name.toLowerCase()).includes(input_username.toLowerCase())) {
                var user = {
                  office: data[i][1],
                  user_name: data[i][2],
                };
                users.push(user);
              }
        }
  }
  

  return users;
}



function findUser(input_office, input_username) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employees'); 
  var data = sheet.getDataRange().getValues();

  
      for (var i = 1; i < data.length; i++) {
      var office = (data[i][1] || '').toString(); 
      var user_name = (data[i][2] || '').toString(); 

      if ( office == input_office && user_name == input_username) {
            var user = {
              office: data[i][1],
              username: data[i][2],
              q1: data[i][3],
              q2: data[i][4],
              q3: data[i][5],
              q4: data[i][6],
              q5: data[i][7],
              q6: data[i][8],
              q7: data[i][9],
              image: data[i][10]
            };
            return user;
          }
    }
  

}
















































