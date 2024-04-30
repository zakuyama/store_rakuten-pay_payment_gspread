const TEST = false

function main() {
  const processStartTime = new Date()

  if(TEST) console.log("=======TEST MODE=======")
  const timestamp = generateTimeStamp(processStartTime)

  const sheetName = TEST ? 'test' : 'paypay'
  const sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName)
  
  const targetMail = getMail()
  // console.log(targetMail)

  //targetMailのメッセージからデータを抽出
  // 1. 楽天キャッシュの使用有無
  //   使用：2に進む
  //   未使用：記録する必要がないので2,3, 4を実行しない
  // 2. 支払日の取得
  // 3. 支払い先の取得
  // 4. スプレッドシートに記載
  // 5. メールを既読にする
  targetMail.map(message =>{
    const plainBody = message.getPlainBody()
    // console.log([plainBody])
    try{
      let rakutenCashValue = getRakutenCash(plainBody)

      if(rakutenCashValue > 0){
        let payDate = getPayDate(plainBody)
        let detail = getPaymentDetail(plainBody)
        console.log(payDate, detail, rakutenCashValue)

        //スプレッドシートに記録
        const record = [timestamp, payDate, rakutenCashValue, '出金', detail, 'paypay', false]
        sheet.appendRow(record)

      }else if(rakutenCashValue < 0){
        throw new Error("Rackten Cash is not negative value.")
      }
      else{
        console.log("No use Rakuten Cash")
      }
      //メールを既読にする
      if(!TEST) GmailApp.markMessageRead(message)

    }catch(e){
        console.log(e)
        console.log([plainBody])
        throw new Error(e)
    }
  })
}

function getMail(){
  const query = 'is:unread from:"no-reply@pay.rakuten.co.jp" subject:"楽天ペイアプリご利用内容確認メール"'
  // is:unread from:(no-reply@pay.rakuten.co.jp) subject:(楽天ペイアプリご利用内容確認メール)
  const start = 0
  const end = 30

  const threads =  GmailApp.search(query, start, end)
  const messgaeForThreads = GmailApp.getMessagesForThreads(threads)

  let targetMails = []

  // 対象とするメールの抽出
  messgaeForThreads.forEach(messgages => {
    // messages : スレッドのデータを1つ取り出したもの
    messgages.forEach(message =>{
      // message : あるスレッドに含まれるメール情報
      // console.log(message)
      // console.log("subject", message.getSubject())
      // console.log("unread", message.isUnread())
      // console.log(message.getPlainBody())
      if(message.isUnread()){
        targetMails.push(message)
      }
    })
  })
  return targetMails
}

function getRakutenCash(body){
  const ptr = /^.+楽天キャッシュ\D*(\d{1,3}|\d{1,3},\d{3})円/m
  const result = body.match(ptr)

  // resultがnullの場合例外を発生させる
  if(!result){
    throw new Error("No match pattern in getRakutenCash")
  }

  let value = result[1]
  value = value.replace(",", "")
  value = parseInt(value)

  return value
}

function getPayDate(body){
  // pattern: ^.+ご利用日時\D*(\d{4}\/\d{2}\/\d{2})
  const ptr = /^.+ご利用日時\D*(\d{4}\/\d{2}\/\d{2})/m

  const result = body.match(ptr)

  // resultがnullの場合例外を発生させる
  if(!result){
    throw new Error("No match pattern in getPayDate")
  }
  //format : 2024/03/18 -> 現金のスプレッドシートには1桁の時の先頭の0がついている。
  return result[1]
}

function getPaymentDetail(body){
  // pattern : (?<=ご利用店舗\W.).+?(?=\W.電話番号)
  // 詳細：https://qiita.com/shotets/items/98f3828b6e5f08d42498
  // const ptr = /(?<=ご利用店舗\W.).+?(?=\W.電話番号)/m
  // 上記パターンでは取得できなかった。下記パターンで何故かマッチした。
  const ptr = /(?<=ご利用店舗\W.).+/m
  const result = body.match(ptr)

  // resultがnullの場合例外を発生させる
  if(!result){
    throw new Error("No match pattern in getPaymentDetail")
  }
  return result[0]
}

function insertSpredSheet(data){

}