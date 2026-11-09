/**
 * taskpane.js — Logika add-in Resize Gambar untuk Word.
 *
 * Versi 1.1 — Memproses:
 *   1. Inline pictures di body dokumen (gambar menyatu dengan teks)
 *   2. Floating shapes di body dokumen (gambar mengambang / wrap text)
 *   3. Inline pictures di header setiap section
 *   4. Inline pictures di footer setiap section
 *   5. Inline pictures di dalam textbox (geometric shape berisi teks+gambar)
 *
 * Cara kerja:
 *   - Word.run() untuk konteks dokumen.
 *   - Untuk tiap scope, kumpulkan kandidat gambar (inline picture & shape),
 *     baca width/height, resize jika lebar > max, tinggi menyesuaikan proporsi.
 *   - Floating shape di Word.js direpresentasikan sebagai `Shape` dengan
 *     `type === "Image"`. Width/height diakses via properti shape.
 *   - Textbox adalah shape dengan `type === "TextBox"`. Isinya berupa
 *     `textFrame.textRange.inlinePictures`. Header/footer juga punya
 *     `inlinePictures` + `shapes` sendiri.
 */

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("error", "Add-in ini hanya berjalan di Word.");
    return;
  }

  document.getElementById("btnResize").addEventListener("click", () => runResize(false));
  document.getElementById("btnPreview").addEventListener("click", () => runResize(true));

  setStatus("idle", 'Siap. Pilih scope, klik <strong>Resize Sekarang</strong> untuk memulai.');
});

/**
 * Baca opsi scope dari checkbox di UI.
 * @returns {{body:boolean, floating:boolean, headerFooter:boolean, textbox:boolean}}
 */
function readScope() {
  return {
    body:        document.getElementById("scope-body").checked,
    floating:    document.getElementById("scope-floating").checked,
    headerFooter:document.getElementById("scope-headerfooter").checked,
    textbox:     document.getElementById("scope-textbox").checked,
  };
}

/**
 * Helper: proses satu koleksi inline pictures.
 * @param {Word.InlinePictureCollection} pics - sudah di-load
 * @param {number} maxLebarPoints
 * @param {number} maxLebarCm
 * @param {boolean} previewOnly
 * @param {string[]} logs
 * @param {string} scopeLabel - prefix label untuk log, misalnya "Body"
 * @returns {number} jumlah gambar yang diubah
 */
function prosesInlinePics(pics, maxLebarPoints, maxLebarCm, previewOnly, logs, scopeLabel) {
  const PT_TO_CM = 2.54 / 72;
  let diubah = 0;

  for (let i = 0; i < pics.items.length; i++) {
    const pic = pics.items[i];
    const lebarAsli = pic.width;
    const tinggiAsli = pic.height;
    const lebarAsliCm = lebarAsli * PT_TO_CM;
    const tinggiAsliCm = tinggiAsli * PT_TO_CM;

    if (lebarAsli > maxLebarPoints) {
      const rasio = maxLebarPoints / lebarAsli;
      const tinggiBaruCm = (tinggiAsli * rasio) * PT_TO_CM;

      if (!previewOnly) {
        pic.width = maxLebarPoints;
        pic.height = tinggiAsli * rasio;
      }
      diubah++;
      logs.push(`[${scopeLabel}] #${i + 1}: ${lebarAsliCm.toFixed(2)} x ${tinggiAsliCm.toFixed(2)} cm -> ${maxLebarCm.toFixed(2)} x ${tinggiBaruCm.toFixed(2)} cm [DIUBAH]`);
    } else {
      logs.push(`[${scopeLabel}] #${i + 1}: ${lebarAsliCm.toFixed(2)} x ${tinggiAsliCm.toFixed(2)} cm [<= ${maxLebarCm} cm, lewati]`);
    }
  }
  return diubah;
}

/**
 * Helper: proses satu koleksi floating shapes yang bertipe Image.
 * @param {Word.ShapeCollection} shapes - sudah di-load
 * @param {number} maxLebarPoints
 * @param {number} maxLebarCm
 * @param {boolean} previewOnly
 * @param {string[]} logs
 * @param {string} scopeLabel
 * @returns {number} jumlah shape yang diubah
 */
function prosesFloatingShapes(shapes, maxLebarPoints, maxLebarCm, previewOnly, logs, scopeLabel) {
  const PT_TO_CM = 2.54 / 72;
  let diubah = 0;

  for (let i = 0; i < shapes.items.length; i++) {
    const sh = shapes.items[i];
    // Hanya proses shape yang BERISI gambar
    // Word.js Shape punya properti .type, .width, .height
    if (sh.type !== "Image") continue;

    const lebarAsli = sh.width;
    const tinggiAsli = sh.height;
    const lebarAsliCm = lebarAsli * PT_TO_CM;
    const tinggiAsliCm = tinggiAsli * PT_TO_CM;

    if (lebarAsli > maxLebarPoints) {
      const rasio = maxLebarPoints / lebarAsli;
      const tinggiBaruCm = (tinggiAsli * rasio) * PT_TO_CM;

      if (!previewOnly) {
        sh.width = maxLebarPoints;
        sh.height = tinggiAsli * rasio;
      }
      diubah++;
      logs.push(`[${scopeLabel} floating] #${i + 1}: ${lebarAsliCm.toFixed(2)} x ${tinggiAsliCm.toFixed(2)} cm -> ${maxLebarCm.toFixed(2)} x ${tinggiBaruCm.toFixed(2)} cm [DIUBAH]`);
    } else {
      logs.push(`[${scopeLabel} floating] #${i + 1}: ${lebarAsliCm.toFixed(2)} x ${tinggiAsliCm.toFixed(2)} cm [<= ${maxLebarCm} cm, lewati]`);
    }
  }
  return diubah;
}

/**
 * Helper: proses gambar di dalam textbox.
 * Word.js: shape.type === "TextBox" punya .textFrame.textRange.inlinePictures
 *
 * @param {Word.ShapeCollection} shapes
 * @param {Word.RequestContext} context
 * @param {number} maxLebarPoints
 * @param {number} maxLebarCm
 * @param {boolean} previewOnly
 * @param {string[]} logs
 * @param {string} scopeLabel
 * @returns {Promise<number>}
 */
async function prosesTextboxPics(shapes, context, maxLebarPoints, maxLebarCm, previewOnly, logs, scopeLabel) {
  let diubah = 0;
  const PT_TO_CM = 2.54 / 72;

  // Kumpulkan textbox shape + load textRange-nya
  const textboxShapes = [];
  for (let i = 0; i < shapes.items.length; i++) {
    const sh = shapes.items[i];
    if (sh.type === "TextBox") {
      textboxShapes.push({ idx: i, shape: sh });
    }
  }

  if (textboxShapes.length === 0) return 0;

  // Load inline pictures di tiap textbox
  const picRefs = [];
  for (const tb of textboxShapes) {
    const pics = tb.shape.textFrame.textRange.inlinePictures;
    pics.load("width, height");
    picRefs.push({ idx: tb.idx, pics });
  }
  await context.sync();

  for (const ref of picRefs) {
    for (let j = 0; j < ref.pics.items.length; j++) {
      const pic = ref.pics.items[j];
      const lebarAsli = pic.width;
      const tinggiAsli = pic.height;
      const lebarAsliCm = lebarAsli * PT_TO_CM;
      const tinggiAsliCm = tinggiAsli * PT_TO_CM;

      if (lebarAsli > maxLebarPoints) {
        const rasio = maxLebarPoints / lebarAsli;
        const tinggiBaruCm = (tinggiAsli * rasio) * PT_TO_CM;

        if (!previewOnly) {
          pic.width = maxLebarPoints;
          pic.height = tinggiAsli * rasio;
        }
        diubah++;
        logs.push(`[${scopeLabel} textbox#${ref.idx + 1}] #${j + 1}: ${lebarAsliCm.toFixed(2)} x ${tinggiAsliCm.toFixed(2)} cm -> ${maxLebarCm.toFixed(2)} x ${tinggiBaruCm.toFixed(2)} cm [DIUBAH]`);
      } else {
        logs.push(`[${scopeLabel} textbox#${ref.idx + 1}] #${j + 1}: ${lebarAsliCm.toFixed(2)} x ${tinggiAsliCm.toFixed(2)} cm [<= ${maxLebarCm} cm, lewati]`);
      }
    }
  }
  return diubah;
}

/**
 * Fungsi utama resize.
 * @param {boolean} previewOnly
 */
async function runResize(previewOnly) {
  const maxLebarCm = parseFloat(document.getElementById("maxWidth").value);
  if (!maxLebarCm || maxLebarCm <= 0) {
    setStatus("error", "Lebar maksimal tidak valid. Masukkan angka > 0.");
    return;
  }

  const scope = readScope();
  if (!scope.body && !scope.floating && !scope.headerFooter && !scope.textbox) {
    setStatus("error", "Pilih minimal satu scope (Body / Floating / Header-Footer / Textbox).");
    return;
  }

  const maxLebarPoints = maxLebarCm * (72 / 2.54);

  const btnResize = document.getElementById("btnResize");
  const btnPreview = document.getElementById("btnPreview");
  btnResize.disabled = true;
  btnPreview.disabled = true;

  setStatus("running", previewOnly
    ? "Menghitung (mode preview, tidak mengubah dokumen)..."
    : "Memproses gambar...");

  try {
    await Word.run(async (context) => {
      const logs = [];
      let totalDiubah = 0;
      let totalDipindai = 0;

      // === 1) BODY: inline pictures ===
      if (scope.body) {
        const bodyPics = context.document.body.inlinePictures;
        bodyPics.load("width, height");
        await context.sync();
        totalDipindai += bodyPics.items.length;
        totalDiubah += prosesInlinePics(bodyPics, maxLebarPoints, maxLebarCm, previewOnly, logs, "Body");
      }

      // === 2) BODY: floating shapes (Image) ===
      if (scope.floating) {
        const bodyShapes = context.document.body.shapes;
        bodyShapes.load("type, width, height");
        await context.sync();
        // Hitung hanya yang Image untuk laporan totalDipindai
        const imgCount = bodyShapes.items.filter(s => s.type === "Image").length;
        totalDipindai += imgCount;
        totalDiubah += prosesFloatingShapes(bodyShapes, maxLebarPoints, maxLebarCm, previewOnly, logs, "Body");
      }

      // === 3) BODY: textbox (geometric shape berisi teks+gambar) ===
      if (scope.textbox) {
        const bodyShapes = context.document.body.shapes;
        bodyShapes.load("type");
        await context.sync();
        const tbCount = bodyShapes.items.filter(s => s.type === "TextBox").length;
        if (tbCount > 0) {
          // totalDipindai dihitung di dalam prosesTextboxPics (per pic)
          const diubah = await prosesTextboxPics(bodyShapes, context, maxLebarPoints, maxLebarCm, previewOnly, logs, "Body");
          totalDiubah += diubah;
        }
      }

      // === 4) HEADER & FOOTER tiap section ===
      if (scope.headerFooter) {
        const sections = context.document.sections;
        sections.load("items");
        await context.sync();

        for (let s = 0; s < sections.items.length; s++) {
          const sec = sections.items[s];
          const header = sec.getHeader(Word.HeaderFooterType.primary);
          const footer = sec.getFooter(Word.HeaderFooterType.primary);
          // Juga firstPage dan evenPage untuk dokumen dengan layout berbeda
          const headerFirst = sec.getHeader(Word.HeaderFooterType.firstPage);
          const footerFirst = sec.getFooter(Word.HeaderFooterType.firstPage);
          const headerEven = sec.getHeader(Word.HeaderFooterType.evenPages);
          const footerEven = sec.getFooter(Word.HeaderFooterType.evenPages);

          // Load inline pictures + shapes dari tiap header/footer
          const targets = [
            { obj: header, label: `Header S${s + 1}` },
            { obj: footer, label: `Footer S${s + 1}` },
            { obj: headerFirst, label: `Header-First S${s + 1}` },
            { obj: footerFirst, label: `Footer-First S${s + 1}` },
            { obj: headerEven, label: `Header-Even S${s + 1}` },
            { obj: footerEven, label: `Footer-Even S${s + 1}` },
          ];

          for (const t of targets) {
            const pics = t.obj.inlinePictures;
            const shs = t.obj.shapes;
            pics.load("width, height");
            shs.load("type, width, height");
          }
          await context.sync();

          for (const t of targets) {
            // inline pictures
            totalDipindai += t.obj.inlinePictures.items.length;
            totalDiubah += prosesInlinePics(t.obj.inlinePictures, maxLebarPoints, maxLebarCm, previewOnly, logs, t.label);
            // floating shapes (Image)
            const imgCount = t.obj.shapes.items.filter(s => s.type === "Image").length;
            totalDipindai += imgCount;
            totalDiubah += prosesFloatingShapes(t.obj.shapes, maxLebarPoints, maxLebarCm, previewOnly, logs, t.label);
            // textbox di header/footer
            if (scope.textbox) {
              const diubah = await prosesTextboxPics(t.obj.shapes, context, maxLebarPoints, maxLebarCm, previewOnly, logs, t.label);
              totalDiubah += diubah;
            }
          }
        }
      }

      // === Sync akhir untuk menerapkan semua perubahan ===
      if (!previewOnly) {
        await context.sync();
      }

      if (totalDipindai === 0 && logs.length === 0) {
        setStatus("idle", "Tidak ada gambar ditemukan di scope yang dipilih.");
        return;
      }

      const modeLabel = previewOnly ? "PREVIEW (tidak mengubah dokumen)" : "TERAPKAN";
      setStatus("success",
        `<strong>Selesai [${modeLabel}]</strong><br/>` +
        `${totalDiubah} gambar akan diubah (dipindai: ${totalDipindai}).<br/>` +
        `<pre>${logs.length > 0 ? logs.join("\n") : "(tidak ada log detail)"}</pre>`
      );
    });
  } catch (err) {
    console.error(err);
    setStatus("error", `Error: ${err.message || err.toString()}<br/><br/>Tip: beberapa API shape mungkin butuh Word API requirement set tertentu. Coba matikan opsi yang bermasalah dan coba lagi.`);
  } finally {
    btnResize.disabled = false;
    btnPreview.disabled = false;
  }
}

/** Helper: ubah kelas & isi div status. */
function setStatus(kind, html) {
  const el = document.getElementById("status");
  el.className = "status status-" + kind;
  el.innerHTML = html;
}
