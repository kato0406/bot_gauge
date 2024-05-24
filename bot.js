const bot = {
  isTarget: false,
  text: '',
  lineBreakCount: 0,
  maxLineBreakCount: 5,
  spreadsheet: SpreadsheetApp.getActiveSpreadsheet(),
  ignoreTitles: [
    'Gauge 脆弱性情報通知サービス',
    'ログイン',
    'マイページ'
  ],
  formattedText(text) {
    return text.replace(' ', '').replace('　', '').replace('\r', '')
  },
  startSearch() {
    bot.text = ''
    bot.lineBreakCount = 0
    bot.isTarget = true
  },
  endSearch() {
    bot.isTarget = false
  },
  setTrigger() {
    const triggers = ScriptApp.getProjectTriggers()
    const date = new Date(dayjs.dayjs().add(1, 'month'))
    date.setHours(1, 30, 0, 0);

    for (const trigger in triggers) {
      ScriptApp.deleteTrigger(triggers[trigger])
    }

    ScriptApp.newTrigger('myFunction')
      .timeBased()
      .at(date)
      .create();

      return
  },
  init() {
    const now = dayjs.dayjs()
    const year = now.year()
    const month = now.format('MM')
    const threads = GmailApp.search(`[Gauge] 脆弱性情報通知メール ${year}/${month}`).reverse()
    const settingSheet = bot.spreadsheet.getSheetByName('基本設定')
    const searchWords = settingSheet.getRange(`A2:A${settingSheet.getLastRow()}`).getValues().flat()
    const inserts = [];
 
    for(thread of threads) {
      const [firstMessage] = thread.getMessages()

      for(text of firstMessage.getBody().split('')) {
        bot.text += text
        
        if(text === '■') {
          bot.startSearch()
        }

        if(!bot.isTarget) continue

        if(text === '\r') {
          bot.lineBreakCount++
        }

        if(bot.lineBreakCount < bot.maxLineBreakCount) continue 

        const [title, url, number, level, status] = bot.text.split('\n')
        const formattedTitle = title.replace('\r', '')
        const checkSearchWord = searchWords.filter((word) => formattedTitle.toUpperCase().includes(word.toUpperCase()))
      
        if (!checkSearchWord.length) continue;

        if (bot.ignoreTitles.includes(formattedTitle)) continue;

        inserts.push([
          [...new Set(checkSearchWord)].join(),
          formattedTitle,
          bot.formattedText(url),
          bot.formattedText(number),
          bot.formattedText(level),
          bot.formattedText(status)
        ])

        bot.endSearch()
      }
    }

    const tagretSheet = bot.spreadsheet.getSheetByName(`${year}${month}`)
    if (tagretSheet) {
      bot.spreadsheet.deleteSheet(tagretSheet)
    }

    const insertSheet = bot.spreadsheet.insertSheet()
    insertSheet.setName(`${year}${month}`).activate()
    bot.spreadsheet.moveActiveSheet(3)

    insertSheet.appendRow(['検索ワード', 'タイトル', 'URL', 'ID', '深刻度', 'ステータス'])
    insertSheet.getRange(`A2:F${inserts.length + 1}`).setValues(inserts.sort())
    insertSheet.getRange(`A1:A${inserts.length + 1}`).createFilter()

    bot.setTrigger()
  }
}

function myFunction() {
  bot.init()
}
