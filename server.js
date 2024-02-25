require("dotenv").config()

const express = require("express")
const line = require("@line/bot-sdk")

// GOOGLE SHEET
const { GoogleSpreadsheet } = require("google-spreadsheet")
const { JWT } = require("google-auth-library")

const config = {
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.CHANNEL_SECRET,
}

const client = new line.messagingApi.MessagingApiClient({
  channelAccessToken: process.env.CHANNEL_ACCESS_TOKEN,
})

// Initialize auth for google sheet
const serviceAccountAuth = new JWT({
  email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
  key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
  scopes: [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
  ],
})

// 關閉編輯的 sheet
const closedSheets = []

async function handleEvent(event) {
  // ignore non-text-message event
  if (event.type !== "message" || event.message.type !== "text") {
    return Promise.resolve(null)
  }

  const userMessage = event.message.text
  const sourceType = event.source.type
  const groupId = event.source?.groupId
  const userId = event.source.userId
  const replyToken = event.replyToken
  try {
    // google sheet
    const doc = new GoogleSpreadsheet(
      process.env.GOOGLE_SHEET_ID,
      serviceAccountAuth
    )
    // Load basic document props and child sheets
    await doc.loadInfo()

    let profile
    if (groupId) {
      profile = await client.getGroupMemberProfile(groupId, userId)
    } else {
      profile = await client.getProfile(userId)
    }

    const { displayName } = profile

    // 判斷是否為管理員
    const isOwner = true

    // 管理員創建新的 sheet
    // "create-D1-A-description-100"
    if (isOwner && userMessage.startsWith("create")) {
      // 檢查訊息格式
      const re = /^create-D[1-5]-[A|B]-[\w\W][^-]+-\d+$/i
      if (!re.test(userMessage))
        throw new Error('請依格式輸入:\n"create-D1-A-描述-單價"')

      const [mode, date, group, description, price] = userMessage
        .toUpperCase()
        .split("-")
      const newSheetTitle = `${date}_${group}`

      // 確認 sheet 是否已存在
      const existSheet = doc.sheetsByTitle[newSheetTitle]
      if (existSheet) throw new Error(`❗工作表 ${newSheetTitle} 已存在❗`)

      // 創建新的 sheet
      const newSheet = await doc.addSheet({
        title: newSheetTitle,
      })

      // 先設定 header row 後才能操作後續的 row
      await newSheet.setHeaderRow(["組合", "單價"], 1)
      await newSheet.setHeaderRow(["id", "名稱", "數量", "價格"], 4)

      await newSheet.loadHeaderRow(1)
      await newSheet.addRow({
        組合: description,
        單價: price,
      })

      return client.replyMessage({
        replyToken,
        messages: [
          {
            type: "text",
            text: `創建 ${newSheetTitle}:${description}`,
          },
        ],
      })
    }

    // 管理員關閉 sheet
    if (isOwner && userMessage.startsWith("close")) {
      const re = /^close-D[1-5]-[A|B]$/i
      if (!re.test(userMessage)) throw new Error('請依格式輸入:\n"close-D1-A"')

      const [mode, date, group] = userMessage.toUpperCase().split("-")
      const sheetTitle = `${date}_${group}`

      // 確認 sheet 是否已關閉
      const isClosed = closedSheets.includes(sheetTitle)
      if (isClosed) throw new Error(`${sheetTitle} 已關閉`)

      closedSheets.push(sheetTitle)
      console.log(closedSheets)

      return client.replyMessage({
        replyToken,
        messages: [
          {
            type: "text",
            text: `已將 ${sheetTitle} 關閉`,
          },
        ],
      })
    }

    // 管理員將關閉的 sheet 開啟
    if (isOwner && userMessage.startsWith("open")) {
      const re = /^open-D[1-5]-[A|B]$/i
      if (!re.test(userMessage)) throw new Error('請依格式輸入:\n"open-D1-A"')

      const [mode, date, group] = userMessage.toUpperCase().split("-")
      const sheetTitle = `${date}_${group}`

      // 確認 sheet 是否已關閉
      const isClosed = closedSheets.includes(sheetTitle)
      if (!isClosed) throw new Error(`${sheetTitle} 已開啟`)

      const idx = closedSheets.indexOf(sheetTitle)
      closedSheets.splice(idx, 1)
      console.log(closedSheets)

      return client.replyMessage({
        replyToken,
        messages: [
          {
            type: "text",
            text: `已將 ${sheetTitle} 開啟`,
          },
        ],
      })
    }

    // 使用者下單
    const orderRe = /^D[1-5]\s*[A|B]\s*[+|-]\s*\d+$/i
    if (orderRe.test(userMessage)) {
      const formatedMessage = userMessage.toUpperCase().replaceAll(" ", "")
      const date = formatedMessage.slice(0, 2)
      const group = formatedMessage.at(2)
      const sheetTitle = `${date}_${group}`
      const amount = +formatedMessage.slice(3)

      if (amount === 0) throw new Error("數量不可為零")

      // 取得指定 sheet
      const sheet = doc.sheetsByTitle[sheetTitle]
      if (!sheet) throw new Error(`${sheetTitle} 尚未開始`)

      // 確認 sheet 是否已關閉
      const isClosed = closedSheets.includes(sheetTitle)
      if (isClosed) throw new Error(`${sheetTitle} 已關閉`)

      // loads range of cells into local cache
      await sheet.loadCells()

      await sheet.loadHeaderRow(4)

      const unitPrice = sheet.getCellByA1("B2").value

      const rows = await sheet.getRows()
      const existRow = rows.find((i) => i.get("id") === userId)
      if (existRow) {
        const originalAmount = +existRow.get("數量")
        const updatedAmount = originalAmount + amount

        if (updatedAmount < 0)
          throw new Error(`數量不可小於零，目前數量 ${originalAmount}`)

        existRow.set("數量", updatedAmount)
        existRow.set("價格", updatedAmount * unitPrice)
        // save changes
        await existRow.save()
      } else {
        // 新增 row
        await sheet.addRow({
          id: userId,
          名稱: displayName,
          數量: amount,
          價格: amount * unitPrice,
        })
      }

      return client.replyMessage({
        replyToken,
        messages: [
          {
            type: "text",
            text: `${displayName}:${date} ${group}${
              amount > 0 ? `+${amount}` : amount
            }`,
          },
        ],
      })
    }
  } catch (err) {
    return client.replyMessage({
      replyToken,
      messages: [
        {
          type: "text",
          text: err.message,
        },
      ],
    })
  }

  return client.replyMessage({
    replyToken,
    messages: [
      {
        type: "text",
        text: event.message.text,
      },
    ],
  })
}

const app = express()

app.post("/webhook", line.middleware(config), (req, res) => {
  Promise.all(req.body.events.map(handleEvent)).then((result) =>
    res.json(result)
  )
})

// listen on port
const port = process.env.PORT || 8080
app.listen(port, () => {
  console.log(`listening on ${port}`)
})
