/**
 * TQ Excel Connector - elevated runner.
 *
 * Deploy with Execute As Role = Administrator. The RESTlet calls this with
 * https.requestSuitelet, so the queries see every posting transaction. Who
 * may ask is checked inside tq_excel_core.js (assertReportAccess).
 *
 * @NApiVersion 2.1
 * @NScriptType Suitelet
 */
define(['./tq_excel_core'], (connector) => {
  const onRequest = (context) => {
    let result;
    try {
      result = connector.handle(JSON.parse(context.request.body || '{}'));
    } catch (e) {
      result = { ok: false, error: e.message || String(e) };
    }
    context.response.setHeader({ name: 'Content-Type', value: 'application/json' });
    context.response.write(JSON.stringify(result));
  };

  return { onRequest };
});
