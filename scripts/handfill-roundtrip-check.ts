/**
 * 手填本 manifest hash 同步機制 round-trip 驗證腳本（一次性，非單元測試）。
 *
 * 跑法：npx tsx scripts/handfill-roundtrip-check.ts
 *
 * 驗證：
 *   1. 沒手改 → hash 相符 → 走 manifest，完整還原（含 manualSort、數字型 customerName）
 *   2. 在 Excel 改過「上」分頁某格 → hash 不符 → 走版面分析，反映改動
 *   3. 改檔情境下，未被改客戶的 manualSort 從 manifest 補回
 */

import ExcelJS from 'exceljs';

import {buildHandfillWorkbook} from '@/writers/handfill-writer';
import {readHandfillBook} from '@/readers/handfill-reader';
import {createEmptyBook, HANDFILL_MANIFEST_SHEET, type HandfillBook} from '@/domain/models/handfill-book';

function makeBook(): HandfillBook {
    const book = createEmptyBook({lineNo: 1, lineName: '彰化', year: 114, month: 5});
    book.id = 'book-fixed';
    book.customers = [
        {
            id: 'cust-A',
            customerId: '12345',
            customerName: '0912345678', // 數字型名稱：測 round-trip 是否被當數字造成偽陽性
            products: [
                {name: '鮮奶', unitPrice: 30},
                {name: '豆漿', unitPrice: 25},
            ],
            restNotes: ['週日休'],
            phones: ['047111222', ''],
            manualSort: true, // 內部欄位：版面表達不了，改檔後須從 manifest 補回
        },
        {
            id: 'cust-B',
            customerId: '678',
            customerName: '王小明',
            // 長品名清單（> 8）測 merge / 區塊撐高
            products: Array.from({length: 10}, (_, i) => ({name: `品項${i + 1}`, unitPrice: 10 + i})),
            restNotes: [],
            phones: ['', ''],
            manualSort: false,
        },
    ];
    return book;
}

async function wbToFile(wb: ExcelJS.Workbook, name: string): Promise<File> {
    const buf = await wb.xlsx.writeBuffer();
    return new File([buf], name, {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
}

/** 把 workbook 序列化、改「上」分頁某格、再序列化成 File（模擬使用者在 Excel 手改）。 */
async function editedFile(wb: ExcelJS.Workbook, name: string, row: number, col: number, value: string): Promise<File> {
    const buf = await wb.xlsx.writeBuffer();
    const wb2 = new ExcelJS.Workbook();
    await wb2.xlsx.load(buf);
    const sheet = wb2.getWorksheet('上');
    if (!sheet) throw new Error('找不到「上」分頁');
    sheet.getCell(row, col).value = value;
    return wbToFile(wb2, name);
}

let failed = 0;

function assert(cond: boolean, msg: string): void {
    if (cond) {
        console.log(`  PASS  ${msg}`);
    } else {
        failed++;
        console.error(`  FAIL  ${msg}`);
    }
}

async function main(): Promise<void> {
    const book = makeBook();

    // ---- Test 1：沒手改 → 走 manifest ----
    console.log('Test 1: 未手改 → 應走 manifest 完整還原');
    const wb1 = await buildHandfillWorkbook(book);
    const file1 = await wbToFile(wb1, 'test.xlsx');
    const r1 = await readHandfillBook(file1);

    // 走 manifest 的證據：id 與 manualSort 與原 book 一致（版面分析會重生 id、manualSort 歸 false）
    assert(r1.id === 'book-fixed', 'book.id 還原（證明走 manifest 而非版面分析）');
    assert(r1.customers[0]?.id === 'cust-A', 'customer.id 還原');
    assert(r1.customers[0]?.manualSort === true, 'manualSort=true 保留');
    assert(r1.customers[0]?.customerName === '0912345678', '數字型 customerName 完整保留（無偽陽性切換）');
    assert(r1.customers[1]?.products.length === 10, '長品名清單完整（10 筆）');

    // ---- Test 2：改「上」分頁第一個品名 → 走版面分析 ----
    console.log('Test 2: 手改品名 → 應走版面分析反映改動');
    // 第一個客戶 block 從 row 5 起（標題 2 + 欄頭 2），品名在 col 2 → B5
    const wb2 = await buildHandfillWorkbook(book);
    const file2 = await editedFile(wb2, 'test.xlsx', 5, 2, '改過的鮮奶');
    const r2 = await readHandfillBook(file2);

    assert(r2.customers[0]?.products[0]?.name === '改過的鮮奶', '手改的品名生效（版面分析）');
    assert(r2.customers[0]?.manualSort === true, '未改客戶的 manualSort 從 manifest 補回（by customerId）');
    assert(r2.customers[0]?.customerName === '0912345678', '其餘欄位仍正確讀出');

    // ---- Test 3：舊格式 manifest（af24e5d，A1 直接是 book、無 layoutHash）----
    console.log('Test 3: 舊格式 manifest → 維持「manifest 優先」舊行為（即使手改也忽略）');
    const wb3 = await buildHandfillWorkbook(book);
    const buf3 = await wb3.xlsx.writeBuffer();
    const wb3b = new ExcelJS.Workbook();
    await wb3b.xlsx.load(buf3);
    // 改寫 manifest 成舊格式（A1 直接 stringify book，無 version/layoutHash 包裝）
    const metaSheet = wb3b.getWorksheet(HANDFILL_MANIFEST_SHEET);
    if (!metaSheet) throw new Error('找不到 manifest 分頁');
    metaSheet.getCell(1, 1).value = JSON.stringify(book);
    // 同時改「上」分頁第一個品名：舊檔無 hash，應被忽略（manifest 優先）
    const upper = wb3b.getWorksheet('上');
    if (!upper) throw new Error('找不到「上」分頁');
    upper.getCell(5, 2).value = '舊檔手改應被忽略';
    const r3 = await readHandfillBook(await wbToFile(wb3b, 'legacy.xlsx'));

    assert(r3.id === 'book-fixed', '舊格式完整還原 book.id');
    assert(r3.customers[0]?.products[0]?.name === '鮮奶', '舊檔手改被忽略（維持 manifest 優先）');
    assert(r3.customers[0]?.manualSort === true, '舊格式 manualSort 仍保留');

    console.log(failed === 0 ? '\nALL PASS' : `\n${failed} FAILED`);
    if (failed > 0) process.exit(1);
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
