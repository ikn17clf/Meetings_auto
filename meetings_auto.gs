const CALENDAR_ID = 'uecrail@gmail.com'; //カレンダーID
const MOTOFILE_ID = 'xxxxxxxxxx' //議事録テンプレートのファイルID
const FOLDER_ID = 'xxxxxxxxxx' //議事録コピー先のフォルダーID

// # 定例会通知 にお知らせ
function postDiscordbot(text) {
  // discord側で作成したボットのウェブフックURL
  const discordWebHookURL = "https://discord.com/api/webhooks/xxxxxxxxxx";

  // 投稿するチャット内容と設定
  const message = {
    "content": text, // チャット本文
    "tts": false  // ロボットによる読み上げ機能を無効化
  }

  const param = {
    "method": "POST",
    "headers": { 'Content-type': "application/json" },
    "payload": JSON.stringify(message)
  }

  UrlFetchApp.fetch(discordWebHookURL, param);
}

// 今日の0時を返す
function getDate0() {
  const date = new Date();
  date.setHours(0);
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}

// 1日前の0時を返す
function getDateBefore1() {
  var date = new Date()
  date.setDate(date.getDate() - 1);
  date.setHours(0)
  date.setMinutes(0)
  date.setSeconds(0)
  date.setMilliseconds(0)
  return date
}

// 3日後の23時59分を返す
function getDateAfter3() {
  var date = new Date()
  date.setDate(date.getDate() + 3);
  date.setHours(23)
  date.setMinutes(59)
  date.setSeconds(0)
  date.setMilliseconds(0)
  return date
}

// eventの日付・曜日・時刻を返す。
function getTime(event){
  var year_month_day = Utilities.formatDate(event.getStartTime(), "Asia/Tokyo", "yyyy/MM/dd").split("/")
  var time1 = Utilities.formatDate(event.getStartTime(), "Asia/Tokyo", "HH:mm")
  var year1 = parseInt(year_month_day[0])
  var month1 = parseInt(year_month_day[1])
  var day1 = parseInt(year_month_day[2])
  var date01 = Moment.moment(String(year1)+"年"+String(month1)+"月"+String(day1)+"日","YYYY年M月D日")
  var dayArray = ["日", "月","火","水","木","金","土"];
  var dayNum = date01.format("d")
  var event_time = [year_month_day, dayArray[dayNum], time1]
  return event_time
}

// eventのdescriptionを整形する。
function event_description_convert(event){
  meeting_contents = event.getDescription()
  if (meeting_contents != ""){
    meeting_contents = meeting_contents.replace(/<br>/g, "\n").replace(/<html-blob>/g, "").replace(/<\/html-blob>/g, "").replace(/<u>/g, "").replace(/<.u>/g, "").replace(/&nbsp;/g, " ").replace(/<a href=".*">/g,"").replace(/<\/a>/g,"")
  }
  return meeting_contents
}

// event_dateの時刻 > now_timeの時刻 のとき、event_date と now_time の差分を秒単位で返す。
function getTimeDiff(event_date, now_time){

  // eventの日時情報
  var year1 = event_date.getFullYear()
  var month1 = event_date.getMonth() + 1
  var day1 = event_date.getDate()
  var hour1 = event_date.getHours()
  var min1 = event_date.getMinutes()
  var sec1 = event_date.getSeconds()
  var date01 = Moment.moment(String(year1)+"年"+String(month1)+"月"+String(day1)+"日","YYYY年M月D日");
  
  // 現在の日時情報
  var year2 = now_time.getFullYear()
  var month2 = now_time.getMonth() + 1
  var day2 = now_time.getDate()
  var hour2 = now_time.getHours()
  var min2 = now_time.getMinutes()
  var sec2 = now_time.getSeconds()
  var date02 = Moment.moment(String(year2)+"年"+String(month2)+"月"+String(day2)+"日","YYYY年M月D日");

  var day_diff = date01.diff(date02,"days");
  var hour_diff = hour1 - hour2
  if (hour_diff < 0){
    hour_diff += 24
    day_diff -= 1
  }
  var min_diff = min1 - min2
  if (min_diff<0){
    min_diff += 60
    hour_diff -= 1
  }
  var sec_diff = sec1 - sec2
  if (sec_diff<0){
    sec_diff += 60
    min_diff -= 1
  }
  return day_diff * 24 * 3600 + hour_diff * 3600 + min_diff * 60 + sec_diff
}

function getFileByName(file_name, folder) {
  var files = DriveApp.getFolderById(folder.getId()).getFilesByName(file_name);
  while (files.hasNext()) {
    // 一つ目のファイルを返す（複数存在した場合は考慮しない）
    return files.next();
  }
  return null;
}

function meeting_support(){
  const calendarId = CALENDAR_ID;
  var today = new Date()
  var cal = CalendarApp.getCalendarById(calendarId)
  var events_list = cal.getEvents(getDateBefore1(), getDateAfter3()) // 1日前から3日後までの予定を取得
  var num_events = events_list.length
  for (var i = 0; i<num_events; ++i){
    var event = events_list[i]
    if (event.getTitle().match("定例会")){ // event に「定例会」という言葉が含まれていたとき
      var start_time_diff = getTimeDiff(event.getStartTime(), today)
      if (start_time_diff < 72*3600 && start_time_diff >= 72*3600-600){ // 3日後に定例会が行われるとき
        // Discord に通知メッセージを送信
        meeting_contents = event_description_convert(event)

        var txts = [
          Utilities.formatDate(event.getStartTime(), "JST", "yyyy/MM/dd"),
          Utilities.formatString("(%s) ", getTime(event)[1]),
          Utilities.formatDate(event.getStartTime(), "JST", "HH:mm"),
          "より定例会を行います．\n",
          Utilities.formatDate(event.getStartTime(), "JST", "yyyy/MM/dd"),
          Utilities.formatString("(%s)", getTime(event)[1]),
          Utilities.formatString("%s\n%s\n", event.getTitle(), meeting_contents)
        ];
        var txt = txts.join("");
        postDiscordbot(txt) 
      }
      else if (start_time_diff < 1800 && start_time_diff >= 1200){ // 定例会までの時間が20分以上30分未満のとき
        meeting_contents = event_description_convert(event)

        var formattedDate = Utilities.formatDate(event.getStartTime(), "JST", "yyyyMMdd");
        var fileName = "電気通信大学鉄道研究会議事録" + "_" + event.getTitle() + "_" + formattedDate;
        //コピー元のファイルを取得します。
        var sourcefile = DriveApp.getFileById(MOTOFILE_ID);
        //コピーを作成します。
        const folder = DriveApp.getFolderById(FOLDER_ID);
        newfile = sourcefile.makeCopy(fileName);

        // メーリングリスト xxxxxxxxxx@googlegroups.com に登録しているgoogleアカウントにファイルの編集権限を付与
        newfile.addEditor('xxxxxxxxxx@googlegroups.com');

        folder.addFile(newfile);

        var txts = [
          Utilities.formatDate(event.getStartTime(), "JST", "HH:mm"),
          "より本日の定例会を始めます．\n",
          Utilities.formatDate(event.getStartTime(), "JST", "yyyy/MM/dd"),
          Utilities.formatString("(%s)", getTime(event)[1]),
          Utilities.formatString("%s\n%s\n", event.getTitle(), meeting_contents),
          "以下のリンクから議事録にアクセスし，「参加者」の自分の学年の所に名前を記入してください．\n",
          Utilities.formatString("議事録のリンク: \n%s", newfile.getUrl()),
        ];
        var txt = txts.join("");
        // Discord に通知メッセージを送信
        postDiscordbot(txt) 
      }
      else if (start_time_diff <= -9000 && start_time_diff > -9600){ // 定例会開始後の時間が150分以上160分未満のとき
        let targetFolder = DriveApp.getFolderById(FOLDER_ID);
        var formattedDate = Utilities.formatDate(event.getStartTime(), "JST", "yyyyMMdd");
        var file_name = Utilities.formatString("電気通信大学鉄道研究会議事録_%s_%s", event.getTitle(), formattedDate)
        var file = getFileByName(file_name, targetFolder)
        var doc = DocumentApp.openByUrl(String(file.getUrl()));
        var paragraphs = doc.getBody().getParagraphs()

        // 鉄研wiki向けに文章を整形する
        var wiki_txts = []
        var wiki_txt = ""
        var num_paragraphs = paragraphs.length
        var hyphen = ""
        for (var i = 0; i<num_paragraphs;++i){
          // インデントの情報を取得。レベル1の箇条書きだと、INDENT_START:36
          // レベル2なら、INDENT_START:72
          // なので、36で割ったときの商を取得して、これを doc_item に入れる。
          // 箇条書きでなければ0、レベル1の箇条書きなら1、レベル2の箇条書きなら2が doc_item に入る。
          var doc_item = paragraphs[i].getAttributes()["INDENT_START"] / 36
          if (doc_item > 3){
            doc_item = 3
          }
          var doc_txt = paragraphs[i].getText()
          for (var j = 0; j<doc_item; ++j){
            hyphen += "-"
          }
          wiki_txts.push(hyphen+doc_txt)
          hyphen = ""
        }
        var event_time = getTime(event)
        var add_sentence = Utilities.formatString("** %s/%s/%s(%s)%s～ ''%s'' [#y%sm%sd%s]", event_time[0][0],event_time[0][1], event_time[0][2], event_time[1], event_time[2], event.getTitle(), event_time[0][0], event_time[0][1], event_time[0][2])
        wiki_txts.unshift(add_sentence)
        wiki_txts[1] = "***" + wiki_txts[1] + Utilities.formatString("[#joi%dm%sd%s]", parseInt(event_time[0][0])-2000,event_time[0][1],event_time[0][2])
        var processing = 0;
        for (processing = 2;!wiki_txts[processing].match("議事録");++processing){
          wiki_txts[processing] += "~"
          if (processing>8){break} // プログラムの暴走を防ぐ
        }
        wiki_txts[processing] += "~"
        for (;processing<num_paragraphs;++processing){
          if (paragraphs[processing].getAttributes()["INDENT_START"] / 36 == 0){
            if (wiki_txts[processing+1] != ""){
              wiki_txts[processing+1] = "''" + wiki_txts[processing+1] + "''"
            }
          }
          if (wiki_txts[processing+1].match("''次回予告''")){break;}
        }
        ++processing
        for (;processing<num_paragraphs;++processing){
          wiki_txts[processing+1] += "~"
        }
        var wiki_txt = wiki_txts.join("\n") // wikiで編集するときに使用する

        var txts1 = [
          "定例会お疲れさまでした．本日の議事録です．"
        ];
        var txts2 = [
          "```",
          Utilities.formatString("%s", wiki_txt),
          "```"
        ];
        var txts3 = [
          "議事録担当者は以下のリンクからwikiにアクセスし，議事録を更新してください．\n",
          "wikiのリンク:\n http://www.uecrail.club.uec.ac.jp/xxxxxxxxxx"
        ];
        var txt1 = txts1.join("");
        var txt2 = txts2.join("");
        var txt3 = txts3.join("");
        // Discord に通知メッセージを送信
        postDiscordbot(txt1)
        postDiscordbot(txt2)
        postDiscordbot(txt3)

        // メーリングリスト xxxxxxxxxx@googlegroups.com に登録しているメールアドレスに定例会の議事録を送る。
        var mail_txts = [
          "定例会お疲れさまでした．本日の議事録です．\n",
          Utilities.formatString("%s\n", wiki_txt),
          "wiki: http://www.uecrail.club.uec.ac.jp/xxxxxxxxxx"
        ];
        var mail_txt = mail_txts.join("\n");
        MailApp.sendEmail({ to: 'xxxxxxxxxx@googlegroups.com', subject: '本日の定例会の議事録',body: mail_txt})
      }
    }
  }
}
