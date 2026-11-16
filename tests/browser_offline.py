import asyncio,json
from pathlib import Path
from playwright.async_api import async_playwright
root=Path(__file__).resolve().parents[1]
async def main():
 async with async_playwright() as p:
  browser=await p.chromium.launch(executable_path='/usr/bin/chromium',args=['--no-sandbox'])
  ctx=await browser.new_context(offline=True)
  page=await ctx.new_page(); errors=[]; requests=[]
  page.on('pageerror',lambda e:errors.append(str(e)))
  page.on('request',lambda r:requests.append(r.url))
  await page.set_content((root/'dist/Redact_v6.3.html').read_text()); await page.wait_for_timeout(700)
  result=await page.evaluate('''async () => {
    const z=new JSZip();z.file('word/document.xml','<root><text>Offline</text></root>');
    const b=await z.generateAsync({type:'arraybuffer'});
    const zip=await RedactLimits.preflight(b);
    return {version:window.ANON_VERSION, title:document.title, zipWorker:zip, fonts:document.fonts.status, buttons:document.querySelectorAll('button').length};
  }''')
  fixture='Контрагент;Сумма\nООО «Вересков»;12000\nАО «Нордлиния»;24000\n'
  await page.locator('#file').set_input_files({'name':'sample.csv','mimeType':'text/csv','buffer':fixture.encode()})
  await page.locator('#step2').wait_for(state='visible')
  await page.locator('#btnAll').click()
  async with page.expect_download() as download_info:
   await page.locator('#btnDownload').click()
  download=await download_info.value; await download.save_as(root/'validation/anonymous.csv')
  red=(root/'validation/anonymous.csv').read_text(encoding='utf-8-sig')
  assert 'Вересков' not in red and 'Нордлиния' not in red
  async with page.expect_download() as key_info:
   await page.locator('#btnKey').click()
  key=await key_info.value; await key.save_as(root/'validation/key.xlsx')
  await page.locator('#tabRestore').click()
  await page.locator('#rFile').set_input_files(root/'validation/anonymous.csv')
  await page.locator('#rFileInfo').filter(has_text='anonymous.csv').wait_for()
  await page.locator('#rKey').set_input_files(root/'validation/key.xlsx')
  await page.locator('#rKeyInfo').filter(has_text='key.xlsx').wait_for()
  async with page.expect_download() as restored_info:
   await page.locator('#btnRestore').click()
  restored=await restored_info.value; await restored.save_as(root/'validation/restored.csv')
  back=(root/'validation/restored.csv').read_text(encoding='utf-8-sig')
  assert back.strip()==fixture.strip(),repr(back)
  result['ui_csv_roundtrip']=True
  # Exercise the updated SheetJS in its real Blob Worker, then the XLS UI roundtrip.
  fixture_xls=root/'tests/fixtures/biff8.xls'
  baseline=await page.evaluate('''async bytes => {
    const converted=await RedactXls.toXlsx(new Uint8Array(bytes).buffer);
    const wb=new ExcelJS.Workbook();await wb.xlsx.load(converted.buffer);
    return wb.worksheets[0].getSheetValues();
  }''',list(fixture_xls.read_bytes()))
  await page.locator('#tabRedact').click()
  await page.locator('#file').set_input_files(fixture_xls)
  await page.locator('#fileInfo').filter(has_text='biff8.xls').wait_for()
  await page.locator('#step2').wait_for(state='visible')
  await page.locator('#btnAll').click()
  # Boolean cells are imported as booleans; keep this column out of textual pseudonymization.
  await page.locator('#colsTable tr[data-c="6"] select[data-k="method"]').select_option('keep')
  async with page.expect_download() as info:
   await page.locator('#btnDownload').click()
  anon=await info.value; await anon.save_as(root/'validation/anonymous.xlsx')
  async with page.expect_download() as info:
   await page.locator('#btnKey').click()
  key=await info.value; await key.save_as(root/'validation/xls-key.xlsx')
  async def workbook_values(path):
   return await page.evaluate('''async bytes => {
    const wb=new ExcelJS.Workbook();await wb.xlsx.load(new Uint8Array(bytes).buffer);
    return wb.worksheets[0].getSheetValues();
   }''',list(path.read_bytes()))
  anon_values=await workbook_values(root/'validation/anonymous.xlsx')
  assert 'Вересков' not in json.dumps(anon_values,ensure_ascii=False,default=str)
  assert 'Нордлиния' not in json.dumps(anon_values,ensure_ascii=False,default=str)
  await page.locator('#tabRestore').click()
  await page.locator('#rFile').set_input_files(root/'validation/anonymous.xlsx')
  await page.locator('#rFileInfo').filter(has_text='anonymous.xlsx').wait_for()
  await page.locator('#rKey').set_input_files(root/'validation/xls-key.xlsx')
  await page.locator('#rKeyInfo').filter(has_text='xls-key.xlsx').wait_for()
  async with page.expect_download() as info:
   await page.locator('#btnRestore').click()
  restored=await info.value; await restored.save_as(root/'validation/restored.xlsx')
  restored_values=await workbook_values(root/'validation/restored.xlsx')
  assert restored_values==baseline,(restored_values,baseline)
  result['ui_xls_roundtrip']=True
  # Legacy byte encodings must also work inside the new full library in a Worker.
  for cp in [1251,866]:
   values=await page.evaluate('''async bytes => {
    const converted=await RedactXls.toXlsx(new Uint8Array(bytes).buffer);
    const wb=new ExcelJS.Workbook();await wb.xlsx.load(converted.buffer);
    return [wb.worksheets[0].getCell('A2').value, wb.worksheets[0].getCell('B2').value];
   }''',list((root/f'tests/fixtures/biff5-cp{cp}.xls').read_bytes()))
   assert values==['ООО "Вересков"','000154'],values
  result['worker_codepages']=[1251,866]
  result.update(errors=errors,network_requests=[x for x in requests if x.startswith(('http:','https:'))],offline=True)
  await page.screenshot(path=str(root/'validation/offline.png'),full_page=True)
  (root/'validation/browser-results.json').write_text(json.dumps(result,ensure_ascii=False,indent=2));print(json.dumps(result,ensure_ascii=False))
  assert not errors and not result['network_requests'] and result['version']=='6.3'
  await browser.close()
asyncio.run(main())
