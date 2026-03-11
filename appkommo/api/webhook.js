const { ConfidentialClientApplication } = require('@azure/msal-node');

// Función maestra: lee los datos así vengan estructurados o aplastados por Kommo
function getFieldValue(body, fieldName, isName = false) {
    if (isName) {
        // Buscar el nombre del cliente
        if (body.leads?.status?.[0]?.name) return body.leads.status[0].name;
        if (body.leads?.update?.[0]?.name) return body.leads.update[0].name;
        if (body.leads?.add?.[0]?.name) return body.leads.add[0].name;
        
        // Buscar en formato aplastado
        for (const key in body) {
            if (key.match(/^leads\[(status|update|add)\]\[0\]\[name\]$/)) {
                return body[key];
            }
        }
        return 'Cliente sin nombre';
    }

    const targetName = fieldName.toLowerCase().trim();

    // 1. Intentar buscar en formato aplastado (El formato rebelde de Kommo)
    for (const key in body) {
        if (key.includes('[custom_fields]') && key.endsWith('[name]')) {
            if (String(body[key]).toLowerCase().trim() === targetName) {
                const basePath = key.substring(0, key.length - 6); // Quita el '[name]'
                const val = body[basePath + '[values][0][value]'];
                const enumVal = body[basePath + '[values][0][enum_code]'];
                const finalVal = (val !== undefined && val !== null && val !== '') ? val : enumVal;
                if (finalVal) return finalVal;
            }
        }
    }

    // 2. Intentar buscar en formato estructurado (Por si un día Kommo lo manda bien)
    const leadsArray = body?.leads?.status || body?.leads?.update || body?.leads?.add || [];
    const lead = Array.isArray(leadsArray) ? leadsArray[0] : Object.values(leadsArray)[0];
    if (lead && lead.custom_fields) {
        const fields = Array.isArray(lead.custom_fields) ? lead.custom_fields : Object.values(lead.custom_fields);
        const field = fields.find(f => f.name && f.name.toLowerCase().trim() === targetName);
        if (field && field.values) {
            const values = Array.isArray(field.values) ? field.values : Object.values(field.values);
            if (values.length > 0) {
                const val = values[0].value;
                const enumVal = values[0].enum_code;
                return (val !== undefined && val !== null && val !== '') ? val : (enumVal || 'N/A');
            }
        }
    }

    return 'N/A';
}

module.exports = async function (req, res) {
    console.log("=== NUEVO DISPARO DE KOMMO ===");
    console.log("DATOS RECIBIDOS:", JSON.stringify(req.body, null, 2));

    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method Not Allowed. Use POST.' });
    }

    try {
        const body = req.body || {};
        
        // Verificamos si al menos viene la palabra "leads" en alguna parte del texto
        const isKommo = Object.keys(body).some(key => key.includes('leads'));
        if (!isKommo) {
            console.error("❌ ERROR: El payload no parece de Kommo.");
            return res.status(400).json({ error: 'Payload irreconocible.' });
        }

        // --- 1. EXTRACCIÓN DE DATOS ---
        const presupuestoRaw = getFieldValue(body, 'Presupuesto');
        const presupuestoNum = parseFloat(presupuestoRaw.toString().replace(/[^0-9.-]+/g, "")) || 0;
        const direccionEntrega = getFieldValue(body, 'Dirección de entrega');
        const tipoBano = getFieldValue(body, 'Tipo de baño');
        const cantidadSanitarios = getFieldValue(body, 'Cantidad de sanitarios');

        const codigo1 = getFieldValue(body, 'Código 1');
        const codigo2 = getFieldValue(body, 'Código 2');
        const codigo3 = getFieldValue(body, 'Código 3');
        const codigo4 = getFieldValue(body, 'Código 4');
        const codigosBano = [codigo1, codigo2, codigo3, codigo4].filter(c => c && c !== 'N/A').join(', ') || 'Sin códigos asignados';

        const contrato = getFieldValue(body, 'No. Contrato');
        const cliente = getFieldValue(body, '', true); // Busca el nombre del cliente
        const contactoEntrega = getFieldValue(body, 'Persona que recibe baño');
        const telefonoEntrega = getFieldValue(body, 'Teléfono persona que recibe');
        const periodoRenta = getFieldValue(body, 'Periodo de renta');

        const notasRaw = getFieldValue(body, 'Notas / Comentarios');
        const notas = notasRaw !== 'N/A' ? notasRaw : getFieldValue(body, 'Notas');
        const metodoPago = getFieldValue(body, 'Método de pago');
        const pagaIvaRaw = getFieldValue(body, 'Paga IVA');
        const direccionPago = getFieldValue(body, 'Dirección de pago');

        // --- 2. LÓGICA CONDICIONAL ---
        const isPagaIva = pagaIvaRaw === true || pagaIvaRaw.toString().toLowerCase() === 'sí' || pagaIvaRaw.toString().toLowerCase() === 'si' || pagaIvaRaw === '1';
        const isEfectivo = metodoPago.toString().toLowerCase().includes('efectivo');

        let saludo = '';
        let costos_html = '';
        let textoSoloEfectivo = '';

        const formatMoney = (val) => `$${val.toLocaleString('es-MX', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

        if (isPagaIva) {
            saludo = "Hola, buen día. ¿Me podrían ayudar a realizar la siguiente facturación y programación, por favor?";
            const subtotal = presupuestoNum;
            const iva = subtotal * 0.16;
            const total = subtotal + iva;
            costos_html = `<strong>Subtotal:</strong> ${formatMoney(subtotal)}<br/><strong>IVA (16%):</strong> ${formatMoney(iva)}<br/><strong>Total:</strong> ${formatMoney(total)}`;
        } else if (!isPagaIva && isEfectivo) {
            saludo = "Hola, buen día. ¿Me podrían ayudar a realizar la siguiente c2020 y programación, por favor?";
            costos_html = `<strong>Presupuesto:</strong> ${formatMoney(presupuestoNum)}`;
            textoSoloEfectivo = `<p style="font-size: 16px; color: #333; margin-top: 5px;"><strong>Dirección de pago:</strong> ${direccionPago}</p>`;
        } else {
            saludo = "Hola, buen día. ¿Me podrían ayudar a realizar la siguiente programación, por favor?";
            costos_html = `<strong>Presupuesto:</strong> ${formatMoney(presupuestoNum)}`;
        }

        // --- 4. FORMATO DE ARCHIVOS ---
        const documentNames = ['CSF', 'Comprobante de domicilio', 'Comprobante de pago', 'INE'];
        let enlacesDocumentos = '';
        documentNames.forEach(doc => {
            const url = getFieldValue(body, doc);
            if (url && (url.startsWith('http') || url.startsWith('www'))) {
                const link = url.startsWith('www') ? `https://${url}` : url;
                enlacesDocumentos += `<li style="margin-bottom: 12px;"><a href="${link}" style="display: inline-block; padding: 10px 16px; background-color: #0078D4; color: #fff; text-decoration: none; border-radius: 6px; font-weight: bold; font-family: sans-serif;">📄 Descargar ${doc}</a></li>`;
            } else if (url && url !== 'N/A') {
                enlacesDocumentos += `<li style="margin-bottom: 10px; font-family: sans-serif;">📄 <strong>${doc}:</strong> ${url}</li>`;
            }
        });

        if (!enlacesDocumentos) {
            enlacesDocumentos = '<li style="font-family: sans-serif;"><em>Sin documentos adjuntos encontrados.</em></li>';
        }

        // --- 3. DISEÑO HTML ---
        const emailHtmlBody = `
        <div style="font-family: Arial, sans-serif; color: #202020; max-width: 650px; margin: 0 auto; outline: 1px solid transparent;">
            <p style="font-size: 16px;">${saludo}</p>
            ${textoSoloEfectivo}
            <table style="width: 100%; border-collapse: collapse; margin-top: 20px; border: 1px solid #ccc; font-size: 15px;">
                <tr><td colspan="2" style="background-color: #c2e5ce; text-align: center; font-weight: bold; padding: 12px; border: 1px solid #ccc; font-size: 18px;">Contrato: ${contrato}</td></tr>
                <tr><td style="padding: 12px; border: 1px solid #ccc; width: 50%; vertical-align: top;"><strong>Cliente:</strong><br/><br/>${cliente}</td><td style="padding: 12px; border: 1px solid #ccc; width: 50%; vertical-align: top;"><strong>Costo:</strong><br/><br/>${costos_html}</td></tr>
                <tr><td style="padding: 12px; border: 1px solid #ccc; vertical-align: top;"><strong>Contacto de entrega:</strong><br/><br/>${contactoEntrega}<br/><a href="tel:${telefonoEntrega}" style="color: #000; text-decoration: none;">${telefonoEntrega}</a></td><td style="padding: 12px; border: 1px solid #ccc; vertical-align: top;"><strong>Código baño:</strong><br/><br/>${codigosBano}</td></tr>
                <tr><td colspan="2" style="background-color: #bce2f3; text-align: center; font-weight: bold; padding: 12px; border: 1px solid #ccc;">Domicilio obra</td></tr>
                <tr><td colspan="2" style="padding: 12px; border: 1px solid #ccc; text-align: center;">${direccionEntrega}</td></tr>
                <tr><td colspan="2" style="background-color: #bce2f3; text-align: center; font-weight: bold; padding: 12px; border: 1px solid #ccc;">Periodo</td></tr>
                <tr><td colspan="2" style="padding: 12px; border: 1px solid #ccc; text-align: center;">${periodoRenta}</td></tr>
                <tr><td colspan="2" style="background-color: #bce2f3; text-align: center; font-weight: bold; padding: 12px; border: 1px solid #ccc;">Descripción</td></tr>
                <tr><td colspan="2" style="padding: 12px; border: 1px solid #ccc; text-align: center;">${cantidadSanitarios} ${tipoBano}</td></tr>
                <tr><td colspan="2" style="background-color: #bce2f3; text-align: center; font-weight: bold; padding: 12px; border: 1px solid #ccc;">Comentarios</td></tr>
                <tr><td colspan="2" style="padding: 15px 12px; border: 1px solid #ccc; text-align: center; white-space: pre-wrap; color: #444;">${notas}</td></tr>
            </table>
            <div style="margin-top: 35px; border-top: 1px solid #eee; padding-top: 15px;">
                <h3 style="color: #333; font-family: Arial, sans-serif; font-size: 18px;">Documentos del Cliente</h3>
                <ul style="list-style-type: none; padding: 0;">${enlacesDocumentos}</ul>
            </div>
        </div>`;

        // --- 5. MICROSOFT GRAPH ---
        if (!process.env.TENANT_ID || !process.env.CLIENT_ID || !process.env.CLIENT_SECRET || !process.env.SENDER_EMAIL) {
            throw new Error("Missing Microsoft API environment variables.");
        }

        const msalConfig = { auth: { clientId: process.env.CLIENT_ID, authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`, clientSecret: process.env.CLIENT_SECRET } };
        const cca = new ConfidentialClientApplication(msalConfig);
        const tokenResponse = await cca.acquireTokenByClientCredential({ scopes: ['https://graph.microsoft.com/.default'] });
        
        // --- DESTINATARIOS --- (Temporalmente solo tú)
        const toRecipientsList = ["d.herrera@saniglobal.com.mx"];

        const sendMailParams = {
            message: { subject: `Nuevo requerimiento - Contrato ${contrato} - Cliente: ${cliente}`, body: { contentType: 'HTML', content: emailHtmlBody }, toRecipients: toRecipientsList.map(email => ({ emailAddress: { address: email } })) },
            saveToSentItems: 'true'
        };

        const graphResponse = await fetch(`https://graph.microsoft.com/v1.0/users/${process.env.SENDER_EMAIL}/sendMail`, {
            method: 'POST', headers: { 'Authorization': `Bearer ${tokenResponse.accessToken}`, 'Content-Type': 'application/json' }, body: JSON.stringify(sendMailParams)
        });

        if (!graphResponse.ok) {
            const errorText = await graphResponse.text();
            console.error("❌ ERROR MICROSOFT GRAPH:", errorText);
            throw new Error(`Graph API Status ${graphResponse.status}: ${errorText}`);
        }

        console.log("✅ ¡CORREO ENVIADO CON ÉXITO!");
        return res.status(200).json({ success: true, message: 'Email procesado y enviado con Microsoft Graph.' });

    } catch (error) {
        console.error("❌ Webhook Falló:", error);
        return res.status(500).json({ error: 'Fallo Interno', details: error.message });
    }
};
