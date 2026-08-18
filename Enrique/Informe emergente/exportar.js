"use strict";
// Construye el libro XLSX de exportación a partir de las filas AGREGADAS
// filtradas ({f, a, s, t, c, n, p}). No depende del DOM: se prueba en node.
(function () {
  const NO_EXITOSA = ["desistida", "cancelada", "calibrador", "rechazada"];
  const CAT_ORDER = ["exitosa", "desistida", "cancelada", "calibrador", "rechazada", "en_proceso", "otros"];
  const CAT_LABEL = {
    exitosa: "Exitosas", desistida: "Cliente desiste", cancelada: "Canceladas",
    calibrador: "Calibrador rechaza", rechazada: "Rechazadas",
    en_proceso: "En proceso", otros: "Otros"
  };

  function porClave(filas, claveFn) {
    const map = new Map();
    for (const r of filas) {
      const k = claveFn(r);
      const cur = map.get(k) || { n: 0, p: 0, porCat: {} };
      cur.n += r.n;
      cur.p += r.p;
      cur.porCat[r.c] = (cur.porCat[r.c] || 0) + r.n;
      map.set(k, cur);
    }
    return map;
  }

  const noDe = (x) => NO_EXITOSA.reduce((s, c) => s + (x.porCat[c] || 0), 0);
  const procDe = (x) => (x.porCat.en_proceso || 0) + (x.porCat.otros || 0);
  const pctDe = (x) => (x.n ? ((x.porCat.exitosa || 0) / x.n) * 100 : 0);

  function hoja(rowsAoa) {
    const ws = XLSX.utils.aoa_to_sheet(rowsAoa);
    ws["!cols"] = rowsAoa[0].map((h) => ({ wch: Math.min(Math.max(String(h).length + 3, 12), 40) }));
    return ws;
  }

  function construirLibro(filas, desde, hasta) {
    const total = filas.reduce((s, r) => s + r.n, 0);
    const resumen = porClave(filas, () => "x").get("x") || { n: 0, p: 0, porCat: {} };
    const exitosas = resumen.porCat.exitosa || 0;
    const valor = filas.reduce((s, r) => s + (r.c === "exitosa" ? r.p : 0), 0);

    const wb = XLSX.utils.book_new();

    // ---- Resumen ----
    const r1 = [
      ["Reporte de Ventas - Sharep"],
      ["Periodo", desde + " → " + hasta],
      [],
      ["Indicador", "Valor"],
      ["Total ventas", total],
      ["Exitosas", exitosas],
      ["% Exito", (total ? (exitosas / total) * 100 : 0).toFixed(1) + "%"],
      ["No exitosas", noDe(resumen)],
      ["En proceso", procDe(resumen)],
      ["Valor planes (exitosas)", Math.round(valor * 100) / 100],
      [],
      ["Ventas por resultado"],
      ["Categoria", "Cantidad", "% del total"],
    ];
    for (const c of CAT_ORDER) {
      const cant = resumen.porCat[c] || 0;
      r1.push([CAT_LABEL[c], cant, (total ? (cant / total) * 100 : 0).toFixed(1) + "%"]);
    }
    r1.push(["TOTAL", total, "100%"]);
    r1.push([]);
    const porTipo = porClave(filas, (r) => r.t);
    r1.push(["Ventas por tipo de venta"], ["Tipo de venta", "Cantidad", "% del total"]);
    for (const [tipo, x] of [...porTipo.entries()].sort((a, b) => b[1].n - a[1].n)) {
      r1.push([tipo, x.n, (x.n / total * 100).toFixed(1) + "%"]);
    }
    XLSX.utils.book_append_sheet(wb, hoja(r1), "Resumen");

    // ---- Diario ----
    const porDia = porClave(filas, (r) => r.f);
    const r2 = [["Fecha", "Total", "Exitosas", "No exitosas", "En proceso"]];
    for (const d of [...porDia.keys()].sort()) {
      const x = porDia.get(d);
      r2.push([d, x.n, x.porCat.exitosa || 0, noDe(x), procDe(x)]);
    }
    XLSX.utils.book_append_sheet(wb, hoja(r2), "Diario");

    // ---- Por mes ----
    const porMes = porClave(filas, (r) => r.f.slice(0, 7));
    const r3 = [["Mes", "Total", "Exitosas", "No exitosas", "En proceso", "% Exito"]];
    for (const m of [...porMes.keys()].sort()) {
      const x = porMes.get(m);
      r3.push([m, x.n, x.porCat.exitosa || 0, noDe(x), procDe(x), pctDe(x).toFixed(1) + "%"]);
    }
    XLSX.utils.book_append_sheet(wb, hoja(r3), "Por mes");

    // ---- Por supervisor ----
    const porSup = porClave(filas, (r) => r.s);
    const r4 = [["Supervisor", "Total", "Exitosas", "No exitosas", "En proceso", "% Exito"]];
    for (const [sup, x] of [...porSup.entries()].sort((a, b) => b[1].n - a[1].n)) {
      r4.push([sup, x.n, x.porCat.exitosa || 0, noDe(x), procDe(x), pctDe(x).toFixed(1) + "%"]);
    }
    XLSX.utils.book_append_sheet(wb, hoja(r4), "Por supervisor");

    // ---- Ranking ----
    const porAs = porClave(filas, (r) => r.a + "|" + r.s);
    const valorAs = new Map();
    for (const r of filas) {
      if (r.c === "exitosa") {
        const k = r.a + "|" + r.s;
        valorAs.set(k, (valorAs.get(k) || 0) + r.p);
      }
    }
    const r5 = [["Asesor", "Supervisor", "Total", "Exitosas", "% Exito", "No exitosas", "En proceso", "Valor planes"]];
    const ranking = [...porAs.entries()]
      .map(([k, x]) => ({ a: k.split("|")[0], s: k.split("|")[1], x }))
      .sort((p, q) => (q.x.porCat.exitosa || 0) - (p.x.porCat.exitosa || 0) || q.x.n - p.x.n);
    for (const { a, s, x } of ranking) {
      const k = a + "|" + s;
      r5.push([a, s, x.n, x.porCat.exitosa || 0, pctDe(x).toFixed(1) + "%", noDe(x), procDe(x), Math.round((valorAs.get(k) || 0) * 100) / 100]);
    }
    XLSX.utils.book_append_sheet(wb, hoja(r5), "Ranking");

    // ---- Datos filtrados (agregados por fecha/asesor/supervisor/tipo/resultado) ----
    const r6 = [["Fecha", "Asesor", "Supervisor", "Tipo de venta", "Resultado", "Cantidad", "Precio total"]];
    for (const r of filas) {
      r6.push([r.f, r.a, r.s, r.t, CAT_LABEL[r.c], r.n, r.p]);
    }
    XLSX.utils.book_append_sheet(wb, hoja(r6), "Datos filtrados");

    return wb;
  }

  window.buildLibroVentas = construirLibro;
})();
