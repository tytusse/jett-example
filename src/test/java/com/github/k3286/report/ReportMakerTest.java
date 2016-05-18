package com.github.k3286.report;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.github.k3286.dto.AnalysisData;

public class ReportMakerTest {

    @Test
    public void report_test() throws Exception {

        // データ準備
        List<AnalysisData> datas = new ArrayList<AnalysisData>();
        datas.add(new AnalysisData(toDate_yyyyMM("201504"), "㈱AVAX", BigDecimal.valueOf(1000000), BigDecimal.valueOf(800000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201505"), "㈱AVAX", BigDecimal.valueOf(1100000), BigDecimal.valueOf(600000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201506"), "㈱AVAX", BigDecimal.valueOf(1300000), BigDecimal.valueOf(600000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201507"), "㈱AVAX", BigDecimal.valueOf(1200000), BigDecimal.valueOf(700000)));

        datas.add(new AnalysisData(toDate_yyyyMM("201504"), "㈱松上電気", BigDecimal.valueOf(4000000), BigDecimal.valueOf(1800000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201505"), "㈱松上電気", BigDecimal.valueOf(3100000), BigDecimal.valueOf(1600000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201506"), "㈱松上電気", BigDecimal.valueOf(2300000), BigDecimal.valueOf(1600000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201507"), "㈱松上電気", BigDecimal.valueOf(3200000), BigDecimal.valueOf(1700000)));

        datas.add(new AnalysisData(toDate_yyyyMM("201504"), "㈱グミシステム", BigDecimal.valueOf(4000000), BigDecimal.valueOf(1800000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201505"), "㈱グミシステム", BigDecimal.valueOf(3100000), BigDecimal.valueOf(1600000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201506"), "㈱グミシステム", BigDecimal.valueOf(2300000), BigDecimal.valueOf(1600000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201507"), "㈱グミシステム", BigDecimal.valueOf(3200000), BigDecimal.valueOf(1700000)));

        datas.add(new AnalysisData(toDate_yyyyMM("201504"), "㈱カフェラテシステム", BigDecimal.valueOf(1000000), BigDecimal.valueOf(600000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201505"), "㈱カフェラテシステム", BigDecimal.valueOf(1000000), BigDecimal.valueOf(600000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201506"), "㈱カフェラテシステム", BigDecimal.valueOf(1000000), BigDecimal.valueOf(600000)));
        datas.add(new AnalysisData(toDate_yyyyMM("201507"), "㈱カフェラテシステム", BigDecimal.valueOf(1000000), BigDecimal.valueOf(600000)));

        // 帳票変換
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("datas", datas);
        Workbook workbook = ReportMaker.toReport(map);

        // ファイル出力
        final String outPath = "output.xlsx";
        try (FileOutputStream fileOut = new FileOutputStream(outPath)) {
            workbook.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public Date toDate_yyyyMM(String dateStr) throws ParseException {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
        return sdf.parse(dateStr);
    }

}
