const SHEET_NAME = 'posts';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet "' + SHEET_NAME + '" not found. Please create it.');
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
