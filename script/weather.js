var API_KEY = PropertiesService.getScriptProperties().getProperty("OPEN_WEATHER_MAP_API_KEY");

// sendEmail はメールで気象情報を送信します。
function sendEmail() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getRange("B2:F32")
  var values = range.getValues();
  for (var i=0; i<values.length; i++) {
    var to = values[i][0];
    var city = values[i][1];
    var country = values[i][2];

    if (to.length < 1) continue;
    var json = getWeatherJSON(city, country)
    Logger.log(json);

    var item = json.list[0];
    var main = item.main;

    var temp = main.temp;
    var pres = main.pressure;
    var humi = main.humidity;
    var tmin = main.temp_min;
    var tmax = main.temp_max;
    var jst = formatTime(new Date(item.dt * 1000))
    var weatherIcon = getWeatherIconString(item.weather[0].id)

    var subject = "[気象情報(日次)] [" + city + "] " + formatTime(new Date());

    var message = "[" + city + "]の気象情報を通知いたします。\n\n";
    message += "## 気象記録日時\n" + jst + "\n\n";
    message += "## 天候\n"
    message += item.weather[0].main + " " + weatherIcon + "\n\n"
    message += "## 気温\n"
    message += "平均気温 " + temp + "度\n";
    message += "最低気温 " + tmin + "度\n";
    message += "最高気温 " + tmax + "度\n\n";
    message += "## 気圧\n" + pres + "hPa\n\n";
    message += "## 湿度\n" + humi + "%\n\n";

    MailApp.sendEmail(to, subject, message)
  }
}

// getWeatherJSON は指定の都市の気象情報を取得します。
function getWeatherJSON(city, country) {
  var url = "http://api.openweathermap.org/data/2.5/find?q=" + city + "," + country + "&units=metric&appid=" + API_KEY;
  var json = UrlFetchApp.fetch(url).getContentText();
  var jsonData = JSON.parse(json);
  return jsonData;
}

// formatTime はDate変数から時刻文字列を生成します。
function formatTime(dt) {
  var year    = dt.getFullYear();
  var month   = ("0" + (dt.getMonth() + 1)).slice(-2);
  var date    = ("0" + dt.getDate()).slice(-2);
  var hours   = ("0" + dt.getHours()).slice(-2);
  var minutes = ("0" + dt.getMinutes()).slice(-2);
  var seconds = ("0" + dt.getSeconds()).slice(-2);
  return year + "/" + month + "/" + date + " " + hours + ":" + minutes + ":" + seconds + " GMT+0900(JST)"
}

// getWeatherIconString は天候コードをアイコン文字列に変換します。
function getWeatherIconString(code) {
  Logger.log(code);
  var weather = Math.floor(code / 100);
  
  var weatherIconString = "";
  switch (weather) {
    case 2:
    case 3:
    case 5:
      weatherIconString = "☔️";
      break;
    case 6:
      weatherIconString = "☃️";
      break;
    case 7:
      weatherIconString = "🌫️";
      break;
    case 8:
      if (code == 800) weatherIconString = "🌞";
      else weatherIconString = "☁️";
      break;
    default:
      weatherIconString = "❓";
      break;
  }
  return weatherIconString;
}
