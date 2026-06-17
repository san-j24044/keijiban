const SHEET_NAME = 'posts';

function setup() {
  getSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName('settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('settings');
    settingsSheet.getRange('A2').setValue('OpenRouter API Key');
    settingsSheet.getRange('B2').setValue('');
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'username', 'category', 'message', 'timestamp', 'likes']);
    sheet.getRange('A1:F1').setFontWeight('bold');
  }
  return sheet;
}

function doGet(e) {
  let response = { success: false };

  try {
    const params = e.parameter || {};
    const mode = params.mode;

    if (mode === 'read') {
      response = readPosts();
    } else if (mode === 'create') {
      response = createPost(params);
    } else if (mode === 'like') {
      response = likePost(params);
    } else if (mode === 'delete') {
      response = deletePost(params);
    } else if (mode === 'simplify') {
      response = simplifyPost(params);
    } else if (mode === 'simplifyText') {
      response = simplifyText(params);
    } else {
      response.error = 'Invalid mode: ' + mode;
    }
  } catch (err) {
    response.success = false;
    response.error = err.toString();
  }

  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  output.setContent(JSON.stringify(response));
  return output;
}

function readPosts() {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { success: true, posts: [] };
  }

  const headers = data[0];
  const posts = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const post = {};
    headers.forEach((header, index) => {
      post[header] = row[index];
    });
    posts.push(post);
  }

  return { success: true, posts: posts };
}

function createPost(params) {
  const username = params.username || '匿名ユーザー';
  const category = params.category || 'general';
  const message = params.message;

  if (!message) {
    return { success: false, error: 'Message is required.' };
  }

  const sheet = getSheet();
  const id = 'p_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
  const timestamp = new Date().toISOString();
  const likes = 0;

  // Column order: id, username, category, message, timestamp, likes
  sheet.appendRow([id, username, category, message, timestamp, likes]);

  return {
    success: true,
    post: { id, username, category, message, timestamp, likes }
  };
}

function likePost(params) {
  const id = params.id;
  if (!id) {
    return { success: false, error: 'ID is required.' };
  }

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const rowIndex = i + 1;
      const currentLikes = Number(data[i][5]) || 0;
      const newLikes = currentLikes + 1;

      sheet.getRange(rowIndex, 6).setValue(newLikes);
      return { success: true, id: id, likes: newLikes };
    }
  }

  return { success: false, error: 'Post with ID ' + id + ' not found.' };
}

function deletePost(params) {
  const id = params.id;
  if (!id) {
    return { success: false, error: 'ID is required.' };
  }

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const rowIndex = i + 1;
      sheet.deleteRow(rowIndex);
      return { success: true, id: id };
    }
  }

  return { success: false, error: 'Post with ID ' + id + ' not found.' };
}

function simplifyPost(params) {
  const id = params.id;
  if (!id) {
    return { success: false, error: 'ID is required.' };
  }

  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  let postMessage = '';

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      postMessage = data[i][3];
      break;
    }
  }

  if (!postMessage) {
    return { success: false, error: 'Post with ID ' + id + ' not found.' };
  }

  return simplifyMessage(postMessage);
}

function simplifyText(params) {
  const message = params.message;

  if (!message) {
    return { success: false, error: 'Message is required.' };
  }

  return simplifyMessage(message);
}

function simplifyMessage(message) {
  const messages = buildSimplifyMessages(message);

  try {
    const simplifiedText = callOpenRouter(messages);
    return { success: true, simplified: simplifiedText };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function buildSimplifyMessages(message) {
  return [
    { role: 'system', content: 'あなたは入力された文章を、小学生でも理解できるような「やさしい日本語」に変換するAIアシスタントです。漢字にはなるべくフリガナ（ひらがな）を付け、難しい言葉は簡単な言葉に言い換えてください。出力は変換後の日本語文章のみにしてください。' },
    { role: 'user', content: '本日の会議は15時から開始します。遅れないようにお願いします。' },
    { role: 'assistant', content: 'きょうのミーティングは、ごご3じからはじまります。おくれないように、じかんまでにあつまってね。' },
    { role: 'user', content: message }
  ];
}

const callOpenRouter = (messages) => {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName('settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('settings');
    settingsSheet.getRange('A2').setValue('OpenRouter API Key');
    settingsSheet.getRange('B2').setValue('');
    throw new Error('スプレッドシートに「settings」シートを自動作成しました。B2セルにOpenRouterのAPIキー（openai/gpt-oss-120b:free 用）を入力してください。');
  }
  const apiKey = settingsSheet.getRange('B2').getValue();
  if (!apiKey) {
    throw new Error('「settings」シートのB2セルにOpenRouterのAPIキーを入力してください。');
  }

  const model = 'openai/gpt-oss-120b:free';

  const url = 'https://openrouter.ai/api/v1/chat/completions';

  const payload = {
    model,
    messages: messages,
    // 任意
    temperature: 0.7,
    max_tokens: 512
  };

  const headers = {
    'Authorization': 'Bearer ' + apiKey,
    'Content-Type': 'application/json',
    'HTTP-Referer': 'https://script.google.com/',
    'X-Title': 'GAS OpenRouter Sample'
  };

  const options = {
    method: 'post',
    headers,
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  const status = res.getResponseCode();
  const text = res.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error('OpenRouter API Error: HTTP ' + status + '\n' + text);
  }

  const json = JSON.parse(text);

  const content = json?.choices?.[0]?.message?.content ?? '';
  Logger.log(content);
  return content;
}
