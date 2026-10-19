/**
 * 外部活頁簿連結（公式裡的 [1]工作表!A1）。
 *
 * ExcelJS 讀檔時會丟掉 xl/externalLinks/ 的連結資料，公式卻原樣寫回；
 * Excel 開到指向不存在連結的公式會判定檔案毀損、無法開啟。
 * 存檔後用原檔把連結資料搬回去，公式與原檔一致；
 * 原檔也沒有連結資料（例如上傳的是已經壞掉的檔）時，那些格改成公式原本算出的數值。
 */

import JSZip from 'jszip';

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EXT_LINK_REL = `${REL_NS}/externalLink`;
const EXT_LINK_CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml';

interface LinkPart {
    xml: string;
    rels: string | null;
}

const attr = (tag: string, name: string): string | undefined =>
    new RegExp(`\\s${name}="([^"]*)"`).exec(tag)?.[1];

/** 相對 xl/ 的 Target 轉成套件內路徑。 */
const partPath = (target: string): string => (target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, ''));

const relsOf = (path: string): string => path.replace(/([^/]+)$/, '_rels/$1.rels');

/** 依 workbook.xml 的 externalReferences 順序（＝公式裡的 [1]、[2]…）讀出各連結資料；讀不到的為 null。 */
async function readLinks(zip: JSZip): Promise<Array<LinkPart | null>> {
    const wb = await zip.file('xl/workbook.xml')?.async('string');
    const rels = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
    const refs = /<externalReferences>([\s\S]*?)<\/externalReferences>/.exec(wb ?? '')?.[1];
    if (!refs || !rels) return [];
    const targets = new Map<string, string>();
    for (const tag of rels.match(/<Relationship\s[^>]*>/g) ?? []) {
        if (attr(tag, 'Type') === EXT_LINK_REL) targets.set(attr(tag, 'Id')!, attr(tag, 'Target')!);
    }
    return Promise.all((refs.match(/<externalReference\s[^>]*>/g) ?? []).map(async (tag) => {
        const target = targets.get(attr(tag, 'r:id') ?? '');
        const path = target && partPath(target);
        const xml = path ? await zip.file(path)?.async('string') : undefined;
        if (!path || !xml) return null;
        return {xml, rels: (await zip.file(relsOf(path))?.async('string')) ?? null};
    }));
}

/** 公式用到的外部活頁簿編號。 */
const extIndexes = (f: string): number[] => [...f.matchAll(/\[(\d+)\]/g)].map((m) => +m[1]);

/** 字串常值裡的 [1] 不算外部參照。 */
const stripStrings = (f: string): string => f.replace(/"(?:[^"]|"")*"/g, '""');

const unescapeXml = (t: string): string => t.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

/**
 * 把 out（ExcelJS 寫出的檔）裡的外部參照補齊：src（使用者上傳的原檔）有的連結資料搬回去，
 * 沒有的把用到它的公式拿掉、保留原本算出的數值。沒有外部參照就原樣回傳。
 */
export async function restoreExternalLinks(out: ArrayBuffer, src: ArrayBuffer): Promise<ArrayBuffer> {
    const zip = await JSZip.loadAsync(out);
    const sheets = Object.keys(zip.files).filter((p) => /^xl\/worksheets\/[^/]+\.xml$/.test(p));
    const xmls = await Promise.all(sheets.map((p) => zip.file(p)!.async('string')));
    const used = new Set<number>();
    for (const xml of xmls) {
        for (const m of xml.matchAll(/<f(?:\s[^>]*)?>([^<]*)<\/f>/g)) extIndexes(stripStrings(unescapeXml(m[1]))).forEach((n) => used.add(n));
    }
    if (!used.size) return out;

    const links = await readLinks(await JSZip.loadAsync(src));
    const ok = (n: number): boolean => !!links[n - 1] && links.slice(0, n).every(Boolean);
    const broken = [...used].filter((n) => !ok(n));

    // 連結資料搬回去（照原本順序，公式裡的編號不用改）
    const keep = links.findIndex((x) => !x) < 0 ? links.length : links.findIndex((x) => !x);
    if (keep && [...used].some(ok)) {
        let rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string');
        let ct = await zip.file('[Content_Types].xml')!.async('string');
        let wb = await zip.file('xl/workbook.xml')!.async('string');
        const refs: string[] = [];
        for (let i = 0; i < keep; i++) {
            const {xml, rels: linkRels} = links[i]!;
            const name = `externalLink${i + 1}.xml`;
            const id = `rIdExtLink${i + 1}`;
            zip.file(`xl/externalLinks/${name}`, xml);
            if (linkRels) zip.file(`xl/externalLinks/_rels/${name}.rels`, linkRels);
            rels = rels.replace('</Relationships>', `<Relationship Id="${id}" Type="${EXT_LINK_REL}" Target="externalLinks/${name}"/></Relationships>`);
            ct = ct.replace('</Types>', `<Override PartName="/xl/externalLinks/${name}" ContentType="${EXT_LINK_CT}"/></Types>`);
            refs.push(`<externalReference r:id="${id}"/>`);
        }
        if (!/xmlns:r=/.test(wb)) wb = wb.replace('<workbook ', `<workbook xmlns:r="${REL_NS}" `);
        // 規格順序：sheets → functionGroups → externalReferences → definedNames
        wb = wb.replace(/<externalReferences>[\s\S]*?<\/externalReferences>/, '')
            .replace(/(<\/sheets>(?:<functionGroups[^>]*\/>|<functionGroups[\s\S]*?<\/functionGroups>)?)/, `$1<externalReferences>${refs.join('')}</externalReferences>`);
        zip.file('xl/_rels/workbook.xml.rels', rels);
        zip.file('[Content_Types].xml', ct);
        zip.file('xl/workbook.xml', wb);
    }

    // 找不到連結資料的公式：拿掉公式，保留原本算出的數值（共用公式連同從屬格一起拿掉）
    if (broken.length) {
        const bad = (f: string): boolean => extIndexes(stripStrings(unescapeXml(f))).some((n) => broken.includes(n));
        sheets.forEach((p, i) => {
            const sis = new Set<string>();
            let xml = xmls[i].replace(/<f(\s[^>]*)?>([^<]*)<\/f>/g, (m, attrs: string | undefined, f: string) => {
                if (!bad(f)) return m;
                const si = attrs && /t="shared"/.test(attrs) ? attr(attrs, 'si') : undefined;
                if (si !== undefined) sis.add(si);
                return '';
            });
            if (sis.size) {
                xml = xml.replace(/<f(\s[^>]*?)\/>/g, (m, attrs: string) =>
                    (/t="shared"/.test(attrs) && sis.has(attr(attrs, 'si') ?? '') ? '' : m));
            }
            if (xml !== xmls[i]) zip.file(p, xml);
        });
    }
    return zip.generateAsync({type: 'arraybuffer', compression: 'DEFLATE'});
}

