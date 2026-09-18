const base = 'http://localhost:3000';
const workbookPath = new URL('../data/HYS_Instagram_Leads.xlsx', import.meta.url);
try {
  const health = await fetch(`${base}/health`, { signal: AbortSignal.timeout(5000) });
  if (health.headers.get('x-hys-local-test') !== 'true') throw new Error('Start the server with LOCAL_TEST_MODE=true before running this script.');
  const payload = {
    object: 'instagram',
    entry: [{ messaging: [{ sender: { id: '123456789' }, message: {
      mid: 'HYS_TEST_001', text: 'Merhaba, 0532 123 45 67 numarasından beni arar mısınız?',
    } }] }],
  };
  const response = await fetch(`${base}/webhook`, {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload), signal: AbortSignal.timeout(30000),
  });
  const responseText = await response.text();
  console.log(`HTTP response: ${response.status} ${response.statusText}`);
  console.log(responseText);
  console.log('Expected normalized phone: +905321234567');
  if (!response.ok) process.exitCode = 1;
  else {
    const result = JSON.parse(responseText);
    const testLead = result.leads?.find(lead => lead.phone === '+905321234567');
    if (!testLead) throw new Error('The response did not contain the expected detected phone.');
    const ExcelJS = (await import('exceljs')).default;
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(workbookPath);
    const worksheet = workbook.getWorksheet('Instagram Leads');
    const actualRow = worksheet?.getRows(2, Math.max(worksheet.rowCount - 1, 0))?.find(row => row.getCell(11).value === 'HYS_TEST_001');
    if (!actualRow || actualRow.getCell(6).value !== '+905321234567') throw new Error('The expected lead is not present in the Excel workbook.');
    console.log(`Excel workbook: ${workbookPath.pathname}`);
    console.log(`Lead row present: yes (inserted: ${testLead.inserted}, reply would be sent: ${testLead.replyWouldBeSent})`);
  }
} catch (error) {
  console.error(error instanceof Error ? error.message : 'Local webhook test failed');
  process.exitCode = 1;
}
