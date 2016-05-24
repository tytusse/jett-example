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

public class InvoiceMakerTest {

    /**
     * 請求書の出力デモ
     * @throws Exception
     */
    @Test
    public void invoice_test() throws Exception {

        Invoice inv = new Invoice();
        inv.setInvoiceNo("INV-00000001");
        inv.setClientName("株式会社 松上電気");
        inv.setAdvancePaid(BigDecimal.valueOf(30000));
        inv.setSalesRep("営業 太郎");
        inv.setInvoiceDate(new Date());

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
