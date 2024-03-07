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

async function handleEvent(event) {
  // ignore non-text-message event
  if (event.type !== "message" || event.message.type !== "text") {
    return Promise.resolve(null)
  }

  const userMessage = event.message.text
  const sourceType = event.source.type // 'user' | 'group'
  const date = new Date(event.timestamp).toLocaleDateString()
  // const groupId = event.source?.groupId
  const groupId = "C1bc0a1b1d91033d5e49aa9b3ca55b751"
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

    const allSheets = doc.sheetsByIndex

    let profile
    if (groupId) {
      profile = await client.getGroupMemberProfile(groupId, userId)
    } else {
      profile = await client.getProfile(userId)
    }

    const { displayName } = profile

    // 判斷是否為管理員
    const isOwner = true

    async function createNewSheet(title, productName, unitPrice, description) {
      // 確認 sheet 是否已存在
      const existSheet = doc.sheetsByTitle[title]
      if (existSheet) throw new Error(`❗${title} 訂單已存在。`)

      // 創建新的 sheet
      const newSheet = await doc.addSheet({
        title,
      })

      // 先設定 header row 後才能操作後續的 row
      await newSheet.setHeaderRow(["狀態", "日期", "商品名", "單價", "描述"], 1)
      await newSheet.setHeaderRow(
        ["id", "客戶名", "數量", "價格", "收款(✅/❌)"],
        4
      )

      await newSheet.loadHeaderRow(1)
      await newSheet.addRow({
        狀態: "開啟",
        日期: date,
        商品名: productName,
        單價: unitPrice,
        描述: description,
      })

      return
    }

    //////////////////////////////////////////////////////////////

    // 管理員操作
    if (isOwner && sourceType === "group") {
      // 創建新訂單 "create-[sheetTitle]-[productName]-[unitPrice]-[description]"
      if (userMessage.startsWith("create")) {
        const re =
          /create-[^-]*[a-zA-Z0-9\u4e00-\u9fa5]+[^-]*-[^-]*[a-zA-Z0-9\u4e00-\u9fa5]+[^-]*-\d+-[^-]*[a-zA-Z0-9\u4e00-\u9fa5]+[^-]*$/
        if (!re.test(userMessage))
          throw new Error("❗創建訂單:\n『 create-訂單名稱-商品名-單價-描述 』")

        const userMessageArr = userMessage.trim().split("-")
        const sheetTitle = userMessageArr[1].toUpperCase().replaceAll(" ", "")
        const productName = userMessageArr[2].trim()
        const unitPrice = userMessageArr[3]
        const description = userMessageArr[4]

        await createNewSheet(sheetTitle, productName, unitPrice, description)

        const replyMessage = `📢 ${productName} 開始登記啦!\n\n${description}\n\n每組價格 ${unitPrice} 元\n\n需要的夥伴請輸入:\t『 ${sheetTitle}+1 』`

        // 發送群組訊息
        // return client.pushMessage({
        //   to: groupId,
        //   messages: [
        //     {
        //       type: "text",
        //       text: replyMessage,
        //     },
        //   ],
        // })
        return client.replyMessage({
          replyToken,
          messages: [
            {
              type: "text",
              text: replyMessage,
            },
          ],
        })
      }

      // 切換訂單狀態 "toggle-[sheetTitle]"
      if (userMessage.startsWith("toggle")) {
        const re = /^toggle\s*-[^-]*[a-zA-Z0-9\u4e00-\u9fa5]+[^-]*$/
        if (!re.test(userMessage))
          throw new Error("❗切換訂單狀態:\n『 toggle-訂單名稱 』")

        const userMessageArr = userMessage.toUpperCase().trim().split("-")
        const sheetTitle = userMessageArr[1].replaceAll(" ", "")

        const sheet = doc.sheetsByTitle[sheetTitle]

        await sheet.loadHeaderRow(1)

        const infoRows = await sheet.getRows({ limit: 1 })

        const sheetStatus = infoRows[0].get("狀態")

        let replyMessage
        if (sheetStatus === "關閉") {
          infoRows[0].set("狀態", "開啟")

          await infoRows[0].save()

          replyMessage = `${sheetTitle} 訂單目前為開啟狀態。`
        }
        if (sheetStatus === "開啟") {
          infoRows[0].set("狀態", "關閉")

          await infoRows[0].save()

          replyMessage = `${sheetTitle} 訂單目前為關閉狀態。`
        }

        return client.replyMessage({
          replyToken,
          messages: [
            {
              type: "text",
              text: replyMessage,
            },
          ],
        })
      }

      // 查看特定使用者的所有訂單 "check-[customer]"
      if (userMessage.startsWith("check")) {
        const re = /^check\s*-[^-]*[a-zA-Z0-9\u4e00-\u9fa5]+[^-]*$/
        if (!re.test(userMessage))
          throw new Error(`❗查詢特定客戶訂單:\n『 check-客戶名稱 』`)

        const customerName = userMessage.trim().split("-")[1]

        let result = ""

        for (const sheet of allSheets) {
          await sheet.loadHeaderRow(1)
          const infoRows = await sheet.getRows({ limit: 1 })

          const sheetTitle = sheet.title
          const productName = infoRows[0]?.get("商品名")
          if (!productName) continue

          await sheet.loadHeaderRow(4)
          const rows = await sheet.getRows({ limit: 30 })
          const existRow = rows.find(
            (row) => row.get("客戶名") === customerName
          )
          if (!existRow) continue

          const amountInSheet = existRow.get("數量")
          const priceInSheet = existRow.get("價格")
          const isCollection = existRow.get("收款") === "✅" ? true : false

          const displayProductName =
            productName.length > 15
              ? `${productName.substring(0, 15)}...`
              : productName

          result += `🛒 [${sheetTitle}]\t\t${
            isCollection ? "已付款" : "未付款"
          }\n\t\t${displayProductName}\n\t\t數量 : ${amountInSheet}\t\t價格 : ${priceInSheet}\n\n`
        }
        if (result.length === 0)
          throw new Error(`❗${customerName} 沒有任何訂單。`)

        const replyMessage = `${customerName} 的所有訂單 :\n${result.slice(
          0,
          -2
        )}`

        return client.replyMessage({
          replyToken,
          messages: [
            {
              type: "text",
              text: replyMessage,
            },
          ],
        })
      }
    }

    //////////////////////////////////////////////////////////////

    //////////////////////////////////////////////////////////////

    // 一般操作
    if (sourceType === "group") {
      const formatedMessage = userMessage.toUpperCase().replaceAll(" ", "")

      // 下單 "[sheetTitle]+1"
      const orderRe = /^\s*[a-zA-Z0-9\u4e00-\u9fa5]+[^-+]*\s*[+|-]\s*\d+\s*$/
      if (orderRe.test(userMessage)) {
        // const [sheetTitle] = formatedMessage.includes("+")
        //   ? formatedMessage.split("+")
        //   : formatedMessage.split("-")
        let sheetTitle, amount
        if (formatedMessage.includes("+")) {
          sheetTitle = formatedMessage.split("+")[0]
          amount = +formatedMessage.split("+")[1]
        }
        if (formatedMessage.includes("-")) {
          sheetTitle = formatedMessage.split("-")[0]
          amount = +formatedMessage.split("-")[1] * -1
        }

        let currentAmount

        if (!amount) throw new Error(`❗${displayName} : 至少要下 1 筆訂單。`)

        // 取得 sheet
        const sheet = doc.sheetsByTitle[sheetTitle]
        if (!sheet)
          throw new Error(`❗${displayName} : ${sheetTitle} 訂單不存在。`)

        const infoRows = await sheet.getRows({ limit: 1 })

        // 確認 sheet 狀態
        const sheetStatus = infoRows[0]?.get("狀態")
        if (sheetStatus !== "開啟")
          throw new Error(`❗${displayName} : ${sheetTitle} 訂單已關閉。`)

        await sheet.loadHeaderRow(4)

        const rows = await sheet.getRows({ limit: 30 })
        const existRow = rows.find((row) => row.get("id") === userId)
        if (existRow) {
          const existRowNumber = existRow.rowNumber
          const amountInSheet = +existRow.get("數量")
          const updatedAmount = amountInSheet + amount

          if (updatedAmount === 0) {
            // 清除 row
            await existRow.delete()

            const replyMessage = `❗${displayName} : 已刪除 ${sheetTitle} 訂單。`

            return client.replyMessage({
              replyToken,
              messages: [
                {
                  type: "text",
                  text: replyMessage,
                },
              ],
            })
          }

          if (updatedAmount < 0)
            throw new Error(
              `❗${displayName} : 數量不可小於零，目前 ${sheetTitle} 訂單數量 ${amountInSheet}。`
            )

          existRow.set("數量", updatedAmount)
          existRow.set("價格", `=C${existRowNumber}*$D$2`)

          // save changes
          await existRow.save()

          currentAmount = updatedAmount
        } else {
          if (amount < 0)
            throw new Error(`❗${displayName} : 至少要下 1 筆訂單。`)

          // 新增 row
          const newRow = await sheet.addRow({
            id: userId,
            客戶名: displayName,
            數量: amount,
            收款: "❌",
          })

          const newRowNumber = newRow.rowNumber

          newRow.set("價格", `=C${newRowNumber}*$D$2`)

          await newRow.save()

          currentAmount = amount
        }

        const replyMessage = `${displayName} :\n${sheetTitle}${
          amount > 0 ? "+" : "-"
        }${Math.abs(amount)}，總共 ${currentAmount} 組。`

        return client.replyMessage({
          replyToken,
          messages: [
            {
              type: "text",
              text: replyMessage,
            },
          ],
        })
      }

      // 查看特定訂單 "[sheetTitle]"
      const specificSheet = doc.sheetsByTitle[formatedMessage]
      if (specificSheet) {
        await specificSheet.loadHeaderRow(4)

        const rows = await specificSheet.getRows({ limit: 30 })

        if (isOwner) {
          let result = ""
          rows.forEach((row, i) => {
            if (!row) return
            const customer = row.get("客戶名")
            const amountInSheet = row.get("數量")
            const isCollection = row.get("收款") === "✅" ? true : false

            result += `${
              i < 9 ? `0${i + 1}` : i + 1
            }. ${customer}+${amountInSheet} ${
              isCollection ? "已付款" : "未付款"
            }\n`
          })

          if (result === "") result = "空空如也..."

          const replyMessage = `${formatedMessage} 訂單 :\n${result.slice(
            0,
            -1
          )}`

          return client.replyMessage({
            replyToken,
            messages: [
              {
                type: "text",
                text: replyMessage,
              },
            ],
          })
        } else {
          const targetRow = rows.find((row) => row.get("id") === userId)
          if (!targetRow)
            throw new Error(`❗${displayName} : 沒有 ${formatedMessage} 訂單。`)

          const amountInSheet = targetRow.get("數量")
          const isCollection = targetRow.get("收款") === "✅" ? true : false

          const replyMessage = `${displayName} :\n🛒 ${formatedMessage}+${amountInSheet}。`
          // const replyMessage = `${displayName} :\n🛒 ${formatedMessage}+${amountInSheet} ${
          //   isCollection ? "已付款" : "未付款"
          // }。`

          return client.replyMessage({
            replyToken,
            messages: [
              {
                type: "text",
                text: replyMessage,
              },
            ],
          })
        }
      }

      // 查看個人所有訂單 "我的訂單"
      if (formatedMessage === "我的訂單") {
        let result = ""
        for (const sheet of allSheets) {
          await sheet.loadHeaderRow(1)
          const infoRows = await sheet.getRows({ limit: 1 })

          const sheetTitle = sheet.title
          const productName = infoRows[0]?.get("商品名")
          if (!productName) continue

          await sheet.loadHeaderRow(4)
          const rows = await sheet.getRows({ limit: 30 })
          const existRow = rows.find((row) => row.get("id") === userId)
          if (!existRow) continue

          const amountInSheet = existRow.get("數量")
          const isCollection = existRow.get("收款") === "✅" ? true : false

          const displayProductName =
            productName.length > 15
              ? `${productName.substring(0, 15)}...`
              : productName

          result += `🛒 [${sheetTitle}] ${displayProductName}\n\t\t數量 : ${amountInSheet}\t\t${
            isCollection ? "已付款" : "未付款"
          }\n\n`
        }
        const replyMessage = `${displayName} 的所有訂單 :\n${result.slice(
          0,
          -2
        )}`

        return client.replyMessage({
          replyToken,
          messages: [
            {
              type: "text",
              text: replyMessage,
            },
          ],
        })
      }
    }

    //////////////////////////////////////////////////////////////
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
