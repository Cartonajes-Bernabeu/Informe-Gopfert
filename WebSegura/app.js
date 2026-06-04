let libroExcelGlobal = null;

async function desbloquearYDesencriptar() {
    const password = document.getElementById('clave').value;
    const errorMsg = document.getElementById('error-msg');
    errorMsg.classList.add('hidden');

    try {
        const response = await fetch('datos.enc.txt');
        if (!response.ok) throw new Error("No se pudo obtener el archivo.");
        const base64Data = await response.text();

        const rawBytes = CryptoJS.enc.Base64.parse(base64Data);
        const ivBytes = CryptoJS.lib.WordArray.create(rawBytes.words.slice(0, 4));
        const encryptedBytes = CryptoJS.lib.WordArray.create(rawBytes.words.slice(4));
        const keyBytes = CryptoJS.enc.Utf8.parse(password.padEnd(32).substring(0, 32));

        const decrypted = CryptoJS.AES.decrypt(
            { ciphertext: encryptedBytes },
            keyBytes,
            { iv: ivBytes, mode: CryptoJS.mode.CBC, padding: CryptoJS.pad.Pkcs7 }
        );

        const libroRaw = decrypted.toString(CryptoJS.enc.Latin1);
        if (!libroRaw || libroRaw.length === 0) throw new Error();

        libroExcelGlobal = XLSX.read(libroRaw, { type: 'binary' });

        const selector = document.getElementById('selector-hojas');
        selector.innerHTML = '';
        libroExcelGlobal.SheetNames.forEach(nombreHoja => {
            const opcion = document.createElement('option');
            opcion.value = nombreHoja;
            opcion.textContent = nombreHoja;
            selector.appendChild(opcion);
        });

        cambiarHoja(libroExcelGlobal.SheetNames[0]);

        document.getElementById('pantalla-bloqueo').classList.add('hidden');
        document.getElementById('contenido-web').classList.remove('hidden');
        document.body.classList.remove('items-center', 'justify-center');

    } catch (error) {
        errorMsg.classList.remove('hidden');
        document.getElementById('clave').value = "";
    }
}

function cambiarHoja(nombreHoja) {
    if (!libroExcelGlobal) return;

    const hoja = libroExcelGlobal.Sheets[nombreHoja];

    // FIX 1: Usar {defval: null} para que las celdas vacías intermedias NO se omitan
    // y las filas conserven su longitud real con índices correctos.
    let datosJSON = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: null });

    // 1. Encontrar la fila de cabecera real (saltando títulos sueltos como "GOPFERT")
    // Una cabecera real tiene al menos 3 celdas con texto no-numérico.
    let indiceCabecera = datosJSON.findIndex(fila => {
        if (!fila) return false;
        const celdasTexto = fila.filter(c => c !== null && c !== "" && typeof c === 'string');
        return celdasTexto.length >= 3;
    });

    if (indiceCabecera === -1) {
        document.getElementById('tabla-cabecera').innerHTML = '<tr><td class="p-4 text-gray-400">No se encontró cabecera</td></tr>';
        document.getElementById('tabla-cuerpo').innerHTML = '';
        return;
    }

    datosJSON = datosJSON.slice(indiceCabecera);

    if (datosJSON.length === 0 || !datosJSON[0]) {
        document.getElementById('tabla-cabecera').innerHTML = '<tr><td class="p-4 text-gray-400">Hoja vacía</td></tr>';
        document.getElementById('tabla-cuerpo').innerHTML = '';
        return;
    }

    // 2. Extraer columnas de la cabecera, recortando vacíos a los extremos
    const cabeceraCruda = datosJSON[0];
    let primerIdx = cabeceraCruda.findIndex(c => c !== null && c !== "");
    if (primerIdx === -1) primerIdx = 0;

    let ultimoIdx = cabeceraCruda.length - 1;
    while (ultimoIdx >= primerIdx && (cabeceraCruda[ultimoIdx] === null || cabeceraCruda[ultimoIdx] === "")) {
        ultimoIdx--;
    }

    // columnas: array de nombres, desde primerIdx hasta ultimoIdx (inclusive)
    const columnas = cabeceraCruda.slice(primerIdx, ultimoIdx + 1);
    const numColumnas = columnas.length;

    // FIX 2: El índice de inicio de datos es el mismo que el de la cabecera (primerIdx).
    // Ya no hay que "buscar" el offset de los datos — con defval:null cada fila tiene
    // los mismos índices que la cabecera, así que simplemente usamos primerIdx.
    const offsetDatos = primerIdx;

    // 3. Dibujar Cabecera con estilo amarillo igual al Excel
    const cabeceraHTML = document.getElementById('tabla-cabecera');
    let filaCabecera = `<tr style="background-color:#FFD966;">`;
    columnas.forEach(col => {
        filaCabecera += `<th style="
            padding: 10px 16px;
            border: 1px solid #c8a400;
            text-transform: uppercase;
            font-size: 0.75rem;
            font-weight: 700;
            color: #1a1a1a;
            white-space: nowrap;
            letter-spacing: 0.05em;
        ">${col !== null && col !== undefined ? col : ""}</th>`;
    });
    filaCabecera += '</tr>';
    cabeceraHTML.innerHTML = filaCabecera;

    // 4. Dibujar Cuerpo de la tabla
    const cuerpoHTML = document.getElementById('tabla-cuerpo');
    let filasCuerpo = '';

    for (let i = 1; i < datosJSON.length; i++) {
        const filaActual = datosJSON[i];
        if (!filaActual) continue;

        // Saltar filas completamente vacías
        if (!filaActual.some(c => c !== null && c !== "")) continue;

        // FIX 3: Recortar la fila exactamente desde el mismo offset que la cabecera
        const filaRecortada = filaActual.slice(offsetDatos, offsetDatos + numColumnas);

        // Detectar si es una fila de totales/global (primera celda contiene "global" o similar)
        const primeracelda = String(filaRecortada[0] || "").toLowerCase();
        const esFilaGlobal = primeracelda.includes('global') || primeracelda.includes('total');

        // Estilo: verde claro para fila Global, igual que en el Excel
        const estiloFila = esFilaGlobal
            ? `style="background-color:#E2EFDA; border-bottom: 1px solid #aed1a0;"`
            : `style="border-bottom: 1px solid rgba(107,114,128,0.3);"`;

        const estiloTd = esFilaGlobal
            ? `style="padding:10px 16px; font-size:0.875rem; white-space:nowrap; color:#1a5e1a; font-weight:600; border: 1px solid #aed1a0;"`
            : `style="padding:10px 16px; font-size:0.875rem; white-space:nowrap; color:inherit; border: 1px solid rgba(107,114,128,0.2);"`;

        filasCuerpo += `<tr ${estiloFila} class="hover-row">`;

        for (let j = 0; j < numColumnas; j++) {
            let valor = filaRecortada[j];

            // FIX 4: Formateo de números — redondear y usar separador de miles español
            // Esto elimina los decimales flotantes (1885794.3329... → 1.885.794)
            if (typeof valor === 'number') {
                // Si el número tiene decimales significativos (no es error de float), los preservamos con 2 decimales
                const redondeado = Math.round(valor * 100) / 100;
                const esEntero = Number.isInteger(redondeado);
                valor = redondeado.toLocaleString('es-ES', {
                    minimumFractionDigits: 0,
                    maximumFractionDigits: esEntero ? 0 : 2
                });
            } else if (valor === null || valor === undefined || valor === "") {
                // FIX 5: Celdas vacías muestran guión "-" igual que en el Excel original
                valor = "-";
            }

            filasCuerpo += `<td ${estiloTd}>${valor}</td>`;
        }

        filasCuerpo += '</tr>';
    }

    cuerpoHTML.innerHTML = filasCuerpo;
}