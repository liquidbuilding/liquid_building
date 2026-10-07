/**
 * 津丘 chill in POS — Google Apps Script 雲端資料庫
 *
 * Google Sheet：
 * https://docs.google.com/spreadsheets/d/18HNOcWckPU9nXjx0mQeif_KpP1_4-jxadrXuSQESc-4/edit
 *
 * Web App：
 * https://script.google.com/macros/s/AKfycbzq6KNy7zxkVRXXWu8_MKWG-LiIv5rOyjAj-tAYTm9Bg-ccTDFSmxgPk5W8w0xhYRQI/exec
 *
 * API Key：
 * iquidbuilding
 *
 * 功能：
 * - 訂單新增 / 修改 / 刪除
 * - 訂單餐點明細
 * - 支出新增 / 修改 / 刪除
 * - 菜單同步
 * - 分類同步
 * - POS 讀取全部雲端資料
 * - ping 測試
 * - JSONP GET 備援
 */

// ============================================================
// 基本設定
// ============================================================

const API_KEY = 'iquidbuilding';

const SPREADSHEET_ID = '18HNOcWckPU9nXjx0mQeif_KpP1_4-jxadrXuSQESc-4';

const SHEET_NAMES = {
  orders: 'Orders',
  items: 'OrderItems',
  expenses: 'Expenses',
  menu: 'Menu',
  categories: 'Categories'
};


// ============================================================
// Google Sheet 欄位
// ============================================================

const HEADERS = {
  Orders: [
    'id',
    'ts',
    'total',
    'subtotal',
    'discount_json',
    'cost',
    'profit',
    'status',
    'paid',
    'updated_at'
  ],

  OrderItems: [
    'order_id',
    'item_index',
    'name',
    'qty',
    'price',
    'cost',
    'note'
  ],

  Expenses: [
    'id',
    'date',
    'name',
    'amount',
    'note',
    'updated_at'
  ],

  Menu: [
    'id',
    'name',
    'unit',
    'price',
    'cost',
    'active',
    'cat',
    'sort_order',
    'discount_type',
    'discount_value',
    'special_price',
    'updated_at'
  ],

  Categories: [
    'id',
    'name',
    'sort_order',
    'updated_at'
  ]
};


// ============================================================
// 建立 / 整理 Google Sheet
// ============================================================

function setupSheets() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  Object.keys(HEADERS).forEach(name => {

    let sh = ss.getSheetByName(name);

    if (!sh) {
      sh = ss.insertSheet(name);
    }

    const headers = HEADERS[name];

    sh.getRange(
      1,
      1,
      1,
      headers.length
    ).setValues([headers]);

    sh.setFrozenRows(1);
  });

  return 'OK';
}


// ============================================================
// JSON 回應
// ============================================================

function json_(obj) {

  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
// JSONP 回應
//
// 給 GitHub Pages / POS 前端做 GET 備援使用。
// 例如：
//
// ?apiKey=iquidbuilding&action=ping&callback=chillinCallback
//
// 回傳：
//
// chillinCallback({"ok":true,...});
//
// ============================================================

function jsonp_(obj, callback) {

  callback = String(callback || '');

  // 只允許正常的 JavaScript callback 名稱
  // 避免把任意內容直接塞進 JavaScript。
  if (
    !/^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callback)
  ) {
    return json_({
      ok: false,
      error: 'JSONP callback 不正確'
    });
  }

  const output =
    callback +
    '(' +
    JSON.stringify(obj) +
    ');';

  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}


// ============================================================
// API Key 驗證
// ============================================================

function checkKey_(key) {

  return String(key || '') === String(API_KEY);

}


// ============================================================
// GET API
//
// 支援：
//
// 1. ping
// ?apiKey=iquidbuilding&action=ping
//
// 2. getData
// ?apiKey=iquidbuilding&action=getData
//
// 3. JSONP
// ?apiKey=iquidbuilding&action=getData&callback=chillinCallback
//
// ============================================================

function doGet(e) {

  try {

    const p =
      e && e.parameter
        ? e.parameter
        : {};

    // ----------------------------------------------------------
    // API Key
    // ----------------------------------------------------------

    if (!checkKey_(p.apiKey)) {

      const errorObj = {
        ok: false,
        error: 'API_KEY 不正確'
      };

      if (p.callback) {
        return jsonp_(errorObj, p.callback);
      }

      return json_(errorObj);
    }


    // ----------------------------------------------------------
    // Action
    // ----------------------------------------------------------

    const action =
      String(p.action || 'ping');


    // ----------------------------------------------------------
    // Ping
    // ----------------------------------------------------------

    if (action === 'ping') {

      const result = {
        ok: true,
        message: '津丘 POS API connected'
      };

      if (p.callback) {
        return jsonp_(result, p.callback);
      }

      return json_(result);
    }


    // ----------------------------------------------------------
    // 取得全部資料
    // ----------------------------------------------------------

    if (action === 'getData') {

      const result = {
        ok: true,
        ...readAll_()
      };

      if (p.callback) {
        return jsonp_(result, p.callback);
      }

      return json_(result);
    }


    // ----------------------------------------------------------
    // 未知 GET action
    // ----------------------------------------------------------

    const errorObj = {
      ok: false,
      error: '未知的 GET action：' + action
    };

    if (p.callback) {
      return jsonp_(errorObj, p.callback);
    }

    return json_(errorObj);


  } catch (err) {

    const errorObj = {
      ok: false,
      error: String(
        err && err.message || err
      )
    };

    if (
      e &&
      e.parameter &&
      e.parameter.callback
    ) {
      return jsonp_(
        errorObj,
        e.parameter.callback
      );
    }

    return json_(errorObj);
  }
}


// ============================================================
// POST API
//
// 支援：
//
// upsertOrder
// deleteOrder
// saveExpense
// deleteExpense
// saveMenu
// saveCategories
//
// ============================================================

function doPost(e) {

  try {

    const raw =
      e &&
      e.postData &&
      e.postData.contents
        ? e.postData.contents
        : '{}';

    const body = JSON.parse(raw);


    // ----------------------------------------------------------
    // API Key
    // ----------------------------------------------------------

    if (!checkKey_(body.apiKey)) {

      return json_({
        ok: false,
        error: 'API_KEY 不正確'
      });

    }


    // ----------------------------------------------------------
    // Action / Payload
    // ----------------------------------------------------------

    const action =
      String(body.action || '');

    const payload =
      body.payload || {};


    // ----------------------------------------------------------
    // Lock
    //
    // 避免手機與電腦同時寫入 Google Sheet
    // 造成資料互相覆蓋。
    // ----------------------------------------------------------

    const lock =
      LockService.getScriptLock();

    lock.waitLock(20000);


    try {

      switch (action) {

        case 'upsertOrder':

          upsertOrder_(
            payload.order
          );

          break;


        case 'deleteOrder':

          deleteOrder_(
            String(
              payload.id || ''
            )
          );

          break;


        case 'saveExpense':

          saveExpense_(
            payload.expense
          );

          break;


        case 'deleteExpense':

          deleteExpense_(
            String(
              payload.id || ''
            )
          );

          break;


        case 'saveMenu':

          replaceMenu_(
            payload.menu || []
          );

          break;


        case 'saveCategories':

          replaceCategories_(
            payload.categories || []
          );

          break;


        default:

          throw new Error(
            '未知的 POST action：' + action
          );

      }

    } finally {

      lock.releaseLock();

    }


    return json_({
      ok: true
    });


  } catch (err) {

    return json_({
      ok: false,
      error: String(
        err && err.message || err
      )
    });

  }
}


// ============================================================
// 取得 Sheet
// ============================================================

function sheet_(name) {

  const ss =
    SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  let sh =
    ss.getSheetByName(name);


  if (!sh) {

    setupSheets();

    sh =
      ss.getSheetByName(name);

  }

  return sh;
}


// ============================================================
// 取得資料列
// ============================================================

function rows_(sh) {

  const last =
    sh.getLastRow();

  const lastCol =
    sh.getLastColumn();


  if (
    last < 2 ||
    lastCol < 1
  ) {
    return [];
  }


  return sh
    .getRange(
      2,
      1,
      last - 1,
      lastCol
    )
    .getValues();
}


// ============================================================
// 用 ID 找資料列
// ============================================================

function findRowById_(sh, id, col) {

  const count =
    Math.max(
      sh.getLastRow() - 1,
      0
    );


  if (count === 0) {
    return -1;
  }


  const vals =
    sh.getRange(
      2,
      col,
      count,
      1
    ).getValues();


  for (
    let i = 0;
    i < vals.length;
    i++
  ) {

    if (
      String(vals[i][0]) ===
      String(id)
    ) {
      return i + 2;
    }

  }


  return -1;
}


// ============================================================
// 訂單：新增 / 修改
// ============================================================

function upsertOrder_(o) {

  if (!o || !o.id) {
    throw new Error(
      '缺少訂單 id'
    );
  }


  const sh =
    sheet_('Orders');


  const now =
    new Date().toISOString();


  const row = [

    String(o.id),

    String(
      o.ts || now
    ),

    Number(o.total) || 0,

    Number(o.subtotal) || 0,

    JSON.stringify(
      o.discount || null
    ),

    Number(o.cost) || 0,

    Number(o.profit) || 0,

    o.status === 'sent'
      ? 'sent'
      : 'pending',

    o.paid
      ? true
      : false,

    now

  ];


  const found =
    findRowById_(
      sh,
      o.id,
      1
    );


  if (found > 0) {

    sh
      .getRange(
        found,
        1,
        1,
        row.length
      )
      .setValues([row]);

  } else {

    sh
      .getRange(
        sh.getLastRow() + 1,
        1,
        1,
        row.length
      )
      .setValues([row]);

  }


  // ----------------------------------------------------------
  // 更新 OrderItems
  // ----------------------------------------------------------

  const itemSh =
    sheet_('OrderItems');


  const existing =
    rows_(itemSh);


  // 先刪除這筆訂單原本的餐點明細
  for (
    let i = existing.length - 1;
    i >= 0;
    i--
  ) {

    if (
      String(existing[i][0]) ===
      String(o.id)
    ) {

      itemSh.deleteRow(i + 2);

    }

  }


  // 再寫入目前最新餐點
  const items =
    Array.isArray(o.items)
      ? o.items
      : [];


  if (items.length) {

    const values =
      items.map(
        (it, i) => [

          String(o.id),

          i,

          String(
            it.name || ''
          ),

          Number(it.qty) || 0,

          Number(it.price) || 0,

          Number(it.cost) || 0,

          String(
            it.note || ''
          )

        ]
      );


    itemSh
      .getRange(
        itemSh.getLastRow() + 1,
        1,
        values.length,
        7
      )
      .setValues(values);

  }

}


// ============================================================
// 訂單：刪除
// ============================================================

function deleteOrder_(id) {

  if (!id) {
    return;
  }


  const sh =
    sheet_('Orders');


  const row =
    findRowById_(
      sh,
      id,
      1
    );


  if (row > 0) {
    sh.deleteRow(row);
  }


  // ----------------------------------------------------------
  // 同時刪除 OrderItems
  // ----------------------------------------------------------

  const itemSh =
    sheet_('OrderItems');


  const existing =
    rows_(itemSh);


  for (
    let i = existing.length - 1;
    i >= 0;
    i--
  ) {

    if (
      String(existing[i][0]) ===
      String(id)
    ) {

      itemSh.deleteRow(
        i + 2
      );

    }

  }

}


// ============================================================
// 支出：新增 / 修改
// ============================================================

function saveExpense_(e) {

  if (!e || !e.id) {

    throw new Error(
      '缺少支出 id'
    );

  }


  const sh =
    sheet_('Expenses');


  const row = [

    String(e.id),

    String(
      e.date || ''
    ),

    String(
      e.name || ''
    ),

    Number(e.amount) || 0,

    String(
      e.note || ''
    ),

    new Date().toISOString()

  ];


  const found =
    findRowById_(
      sh,
      e.id,
      1
    );


  if (found > 0) {

    sh
      .getRange(
        found,
        1,
        1,
        row.length
      )
      .setValues([row]);

  } else {

    sh
      .getRange(
        sh.getLastRow() + 1,
        1,
        1,
        row.length
      )
      .setValues([row]);

  }

}


// ============================================================
// 支出：刪除
// ============================================================

function deleteExpense_(id) {

  if (!id) {
    return;
  }


  const sh =
    sheet_('Expenses');


  const row =
    findRowById_(
      sh,
      id,
      1
    );


  if (row > 0) {
    sh.deleteRow(row);
  }

}


// ============================================================
// 菜單：整批替換
// ============================================================

function replaceMenu_(menu) {

  const sh =
    sheet_('Menu');


  const clean =
    Array.isArray(menu)
      ? menu
      : [];


  const last =
    sh.getLastRow();


  if (last > 1) {

    sh
      .getRange(
        2,
        1,
        last - 1,
        HEADERS.Menu.length
      )
      .clearContent();

  }


  if (!clean.length) {
    return;
  }


  const now =
    new Date().toISOString();


  const values =
    clean.map(
      (m, i) => [

        String(
          m.id || ''
        ),

        String(
          m.name || ''
        ),

        String(
          m.unit || ''
        ),

        Number(m.price) || 0,

        Number(m.cost) || 0,

        m.active !== false,

        String(
          m.cat || ''
        ),

        Number(
          m.order
        ) || i,

        m.saleDiscount
          ? String(
              m.saleDiscount.type ||
              'percent'
            )
          : '',

        m.saleDiscount
          ? Number(
              m.saleDiscount.value
            ) || 0
          : '',

        Number(
          m.specialPrice
        ) || 0,

        now

      ]
    );


  sh
    .getRange(
      2,
      1,
      values.length,
      HEADERS.Menu.length
    )
    .setValues(values);

}


// ============================================================
// 分類：整批替換
// ============================================================

function replaceCategories_(categories) {

  const sh =
    sheet_('Categories');


  const clean =
    Array.isArray(categories)
      ? categories
      : [];


  const last =
    sh.getLastRow();


  if (last > 1) {

    sh
      .getRange(
        2,
        1,
        last - 1,
        HEADERS.Categories.length
      )
      .clearContent();

  }


  if (!clean.length) {
    return;
  }


  const now =
    new Date().toISOString();


  const values =
    clean.map(
      (c, i) => [

        String(
          c.id || ''
        ),

        String(
          c.name || ''
        ),

        i,

        now

      ]
    );


  sh
    .getRange(
      2,
      1,
      values.length,
      4
    )
    .setValues(values);

}


// ============================================================
// 讀取全部雲端資料
// ============================================================

function readAll_() {

  const ordersSh =
    sheet_('Orders');

  const itemSh =
    sheet_('OrderItems');

  const expenseSh =
    sheet_('Expenses');

  const menuSh =
    sheet_('Menu');

  const catSh =
    sheet_('Categories');


  // ----------------------------------------------------------
  // OrderItems → itemMap
  // ----------------------------------------------------------

  const itemMap = {};


  rows_(itemSh).forEach(
    r => {

      const id =
        String(
          r[0] || ''
        );


      if (!itemMap[id]) {
        itemMap[id] = [];
      }


      itemMap[id].push({

        name:
          String(
            r[2] || ''
          ),

        qty:
          Number(r[3]) || 0,

        price:
          Number(r[4]) || 0,

        cost:
          Number(r[5]) || 0,

        note:
          String(
            r[6] || ''
          )

      });

    }
  );


  // ----------------------------------------------------------
  // Orders
  // ----------------------------------------------------------

  const orders =
    rows_(ordersSh)

      .map(
        r => ({

          id:
            String(
              r[0] || ''
            ),

          ts:
            String(
              r[1] || ''
            ),

          total:
            Number(r[2]) || 0,

          subtotal:
            Number(r[3]) || 0,

          discount:
            parseJson_(
              r[4],
              null
            ),

          cost:
            Number(r[5]) || 0,

          profit:
            Number(r[6]) || 0,

          status:
            String(r[7]) === 'sent'
              ? 'sent'
              : 'pending',

          paid:
            r[8] === true ||
            String(
              r[8]
            ).toLowerCase() ===
              'true',

          items:
            itemMap[
              String(
                r[0] || ''
              )
            ] || []

        })
      )

      // 保留有餐點明細的訂單
      .filter(
        o =>
          o.id &&
          o.items.length
      );


  // ----------------------------------------------------------
  // Expenses
  // ----------------------------------------------------------

  const expenses = {};


  rows_(expenseSh).forEach(
    r => {

      const date =
        String(
          r[1] || ''
        );


      if (!expenses[date]) {
        expenses[date] = [];
      }


      expenses[date].push({

        id:
          String(
            r[0] || ''
          ),

        name:
          String(
            r[2] || ''
          ),

        amount:
          Number(r[3]) || 0,

        note:
          String(
            r[4] || ''
          )

      });

    }
  );


  // ----------------------------------------------------------
  // Menu
  // ----------------------------------------------------------

  const menu =
    rows_(menuSh)

      .map(
        r => ({

          id:
            String(
              r[0] || ''
            ),

          name:
            String(
              r[1] || ''
            ),

          unit:
            String(
              r[2] || ''
            ),

          price:
            Number(r[3]) || 0,

          cost:
            Number(r[4]) || 0,

          active:
            r[5] !== false &&
            String(
              r[5]
            ).toLowerCase() !==
              'false',

          cat:
            String(
              r[6] || ''
            ),

          order:
            Number(r[7]) || 0,

          saleDiscount:
            r[8]
              ? {
                  type:
                    String(r[8]) ===
                    'fixed'
                      ? 'fixed'
                      : 'percent',

                  value:
                    Number(r[9]) || 0
                }
              : null,

          specialPrice:
            Number(r[10]) || 0

        })
      )

      .filter(
        m =>
          m.id &&
          m.name
      );


  // ----------------------------------------------------------
  // Categories
  // ----------------------------------------------------------

  const categories =
    rows_(catSh)

      .map(
        r => ({

          id:
            String(
              r[0] || ''
            ),

          name:
            String(
              r[1] || ''
            ),

          order:
            Number(r[2]) || 0

        })
      )

      .filter(
        c =>
          c.id &&
          c.name
      )

      .sort(
        (a, b) =>
          a.order - b.order
      )

      .map(
        c => ({

          id: c.id,

          name: c.name

        })
      );


  // ----------------------------------------------------------
  // 回傳
  // ----------------------------------------------------------

  return {

    orders,

    expenses,

    menu,

    categories

  };

}


// ============================================================
// 安全解析 JSON
// ============================================================

function parseJson_(v, fallback) {

  try {

    return v
      ? JSON.parse(
          String(v)
        )
      : fallback;

  } catch (e) {

    return fallback;

  }

}
