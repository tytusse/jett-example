package com.github.k3286.report;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.github.k3286.dto.Invoice;
import com.github.k3286.dto.InvoiceDetail;

public class InvoiceMakerTest {

    private static BigDecimal TAX_RATE = new BigDecimal(0.08);

    /**
     * 請求書の出力デモ
     * @throws Exception
     */
    @Test
    public void invoice_test() throws Exception {

        Invoice inv = new Invoice();
        inv.setInvoiceNo("INV-00000001");
        inv.setClientPostCode("〒123-3333");
        inv.setClientAddress("東京都品川区東五反田１丁目６−３ 東京建物五反田ビル 108F");
        inv.setClientName("株式会社 松上電気");
        inv.setSalesRep("営業 太郎");
        inv.setInvoiceDate(new Date());

        // 明細行は5行にしておく
        for (int idx = 1; idx <= 5; idx++) {
            InvoiceDetail dtl = new InvoiceDetail();
            dtl.setItemName("サンプル明細ですよ " + idx);
            dtl.setUnitCost(BigDecimal.valueOf(10000));
            dtl.setQuantity(Double.valueOf(idx));
            dtl.setAmt(dtl.getUnitCost().multiply(//
                    BigDecimal.valueOf(dtl.getQuantity())));
            inv.getDetails().add(dtl);
        }
        BigDecimal total = BigDecimal.ZERO;
        for (InvoiceDetail dtl : inv.getDetails()) {
            total = total.add(dtl.getAmt());
        }
        // 立替金
        inv.setAdvancePaid(BigDecimal.valueOf(10800));
        // 税額
        inv.setTaxAmt(total.multiply(TAX_RATE));
        // 請求額（税込）
        inv.setInvoiceAmtTaxin(total.add(inv.getTaxAmt()).add(inv.getAdvancePaid()));
        // 備考
        inv.setNote("これは備考です、サンプルとして備考を記述し、"
                + "そして帳票に出力をしてみました。"
                + "折り返してくれるといいのですが、どうでしょうか");

        // 帳票変換
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("inv", inv);

        Workbook workbook = ReportMaker.toReport(map, "template_invoice.xlsx");

        // ファイル出力
        final String outPath = "output_invoice.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(outPath)) {
            workbook.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
