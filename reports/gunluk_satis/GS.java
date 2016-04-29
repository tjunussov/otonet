// New version created
// 118  shift to start BIRIM_TIP_CODE = 03
// MAGAZA TOPLAM shift to start 
// Pages LFL (BireBir),and TOTALS все магазины против активных
// 01 - LFL, 02 <> General, BIRIM_TIP_KOD = 02 IN (100TS,200TS),  03 - Sanal
// BIRIM_DURUM_KOD
// LFL = BIRIM_TIP_KO = 01, RAPOR_KOD=01
  	  
  // +Gecen yil MUSTERI SAYSI column, toplam, 
// ++gecen yil percentage diff column

package kz.ramportal.reports;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Arrays;

import kz.bee.wx.properties.PropertyManager;
import kz.bee.wx.properties.PropertyMap;
import oracle.jdbc.driver.OracleTypes;

import org.apache.poi.hpsf.WritingNotSupportedException;
import org.apache.poi.hssf.record.CFRuleRecord.ComparisonOperator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFConditionalFormattingRule;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFFontFormatting;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSheetConditionalFormatting;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.jboss.seam.annotations.Logger;
import org.jboss.seam.annotations.Name;
import org.jboss.seam.log.Log;

@Name("GSReport")
public class GSReport implements Report {

	@Logger
	Log log;

	public void run() {
		try {

			String[] tmpStoreCodeList = { "100", "101", "103", "104", "105", "200", "201", "300", "400", "600" };
			byte[] data = execute(new java.util.Date(), new java.util.Date(), tmpStoreCodeList);

			FileOutputStream fos = new FileOutputStream("GS.xls");
			fos.write(data);
			fos.close();

		} catch (IllegalAccessException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public byte[] execute(java.util.Date dateFrom, java.util.Date donem, String[] storeList) throws Exception, IllegalAccessException, ClassNotFoundException,
			SQLException {

		PropertyMap properties = PropertyManager.instance().get("D_ITDB_CONNECTION");

		String url = properties.get("URL");
		String username = properties.get("USERNAME");
		String password = properties.get("PASSWORD");

		return execute(url, username, password, donem, storeList);
	}

	Connection conn;
	java.util.Date donem = null;
	String[] storeList;

	public byte[] execute(String url, String username, String password, java.util.Date donem, String[] storeList) throws Exception {

		try {
			this.donem = donem;

			this.storeList = storeList;
			Arrays.sort(this.storeList);

			Class.forName("oracle.jdbc.driver.OracleDriver").newInstance();
			conn = DriverManager.getConnection(url, username, password);

			log.info("report4.execute");

			wb = new HSSFWorkbook();
			
			

			createFirstSheet();
			createSecondSheet();

			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			wb.write(baos);

			return baos.toByteArray();
		} catch (Exception e) {
			throw e;
		} finally {
			conn.close();
		}

	}

	HSSFWorkbook wb;
	HSSFSheet sheet2;
	HSSFSheet sheet1;

	HSSFCellStyle style1 = null, style1left = null, style2simple = null, style2perc = null, style4simple = null, style4perc = null, style8Text = null;
	HSSFCellStyle style2 = null;
	HSSFCellStyle style2centered = null;
	HSSFCellStyle style4 = null;

	short pos;
	
	ArrayList<Region> regions = null;

	private void createSecondSheet() throws SQLException, IOException, WritingNotSupportedException {

		sheet2 = wb.createSheet("USD");
		sheet2.createFreezePane(1, 4);

		initStyle();
		
		HSSFRow header = sheet2.createRow((short) 2);
		header.setHeight((short) 350);

		HSSFRow subheader = sheet2.createRow((short) 3);
		subheader.setHeight((short) 700);

		HSSFRow title = sheet2.createRow((short) 0);
		createCell(title, 0, "RAMSTOR KAZAKISTAN", style1);

		addMergedRegion(sheet2, wb, 2, 0, 3, 0);
		createCell(header, 0, "TARIH", style1);
		setColumnWidth(sheet2, 0, 5000);

		// -------- Drawing in width magazas

		megaTotalColumns[0] = "";
		megaTotalColumns[1] = "";
		megaTotalColumns[2] = "";
		megaTotalColumns[3] = "";
		
		regions = new ArrayList<Region>();

		int index = 0;

		for (String store : storeList) {

			CallableStatement cstmt1 = conn.prepareCall("call RAM_REPORTS.gsusd(?,?,?)");
			cstmt1.registerOutParameter(1, OracleTypes.CURSOR);

			cstmt1.setString(2, store);
			cstmt1.setDate(3, new Date(donem.getTime()));
			cstmt1.execute();

			pos = drawMagazaDetails(store, index++, sheet2, header, subheader, (ResultSet) cstmt1.getObject(1));

			cstmt1.close();

		}

		// -------- Drawing in magazas toplam details, pivot magaza 100

		{
			CallableStatement cstmt1t = conn.prepareCall("call RAM_REPORTS.gsusd(?,?,?)");
			cstmt1t.registerOutParameter(1, OracleTypes.CURSOR);

			cstmt1t.setString(2, "100");
			cstmt1t.setDate(3, new Date(donem.getTime()));
			cstmt1t.execute();

			pos = drawMagazaDetailsToplams(index++, sheet2, header, subheader, (ResultSet) cstmt1t.getObject(1));

			cstmt1t.close();
		}

		// -------- Drawing magaza toplam excluding toptan satis
		{

			CallableStatement cstmt2 = conn.prepareCall("call RAM_REPORTS.gsmtusd(?,?)");
			cstmt2.registerOutParameter(1, OracleTypes.CURSOR);

			cstmt2.setDate(2, new Date(donem.getTime()));
			cstmt2.execute();

			pos = drawMagazaToplams(sheet2, (ResultSet) cstmt2.getObject(1));

			cstmt2.close();

		}

		// -------- Drawing magaza toplam only toptan satis
		{

			CallableStatement cstmt3 = conn.prepareCall("call RAM_REPORTS.gstsusd(?,?)");
			cstmt3.registerOutParameter(1, OracleTypes.CURSOR);

			cstmt3.setDate(2, new Date(donem.getTime()));
			cstmt3.execute();

			pos = drawToptanSatis(sheet2, (ResultSet) cstmt3.getObject(1));

			cstmt3.close();

		}

		// -------- Drawing megatoplams
		{

			CallableStatement cstmt4 = conn.prepareCall("call RAM_REPORTS.gstsusd(?,?)");
			cstmt4.registerOutParameter(1, OracleTypes.CURSOR);

			cstmt4.setDate(2, new Date(donem.getTime()));
			cstmt4.execute();

			pos = drawSirketToplams(sheet2, (ResultSet) cstmt4.getObject(1));

			cstmt4.close();

		}
		
		HSSFSheetConditionalFormatting cf = sheet2.getSheetConditionalFormatting();

		HSSFConditionalFormattingRule rule = cf.createConditionalFormattingRule(ComparisonOperator.LT, "0",null);
		HSSFFontFormatting fntFrm = rule.createFontFormatting();
		fntFrm.setFontColorIndex(HSSFFont.COLOR_RED);
		
		Region[] regs = new Region[regions.size()];
		regions.toArray(regs);
		
		cf.addConditionalFormatting(regs, rule);

	}

	private short drawSirketToplams(HSSFSheet sheet, ResultSet rset) throws SQLException {

		pos += 2;

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row1 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row2 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row3 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row4 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row5 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row6 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row7 = sheet.createRow((short) pos++);

		short c = 0;

		createCell(row1, c, "SIRKET TOPLAM", style1);
		createCell(row2, c, "NET SATIŞ - GEÇEN YIL", style1left);
		createCell(row3, c, "NET SATIŞ - BÜTÇE", style1left);
		createCell(row4, c, "NET SATIS -FİİLİ", style1left);
		createCell(row5, c, "Prog.Göre Artış", style1left);
		createCell(row6, c, "G.Yıla Göre Artış", style1left);
		createCell(row7, c, "MUSTERI SAYISI", style1left);

		String musteriRow = sirketToplamColums[3].split(",")[0];

		while (rset.next()) {

			c++;

			createCell(row1, c, rset.getString(1), style1);
			createFormualCell(row2, c, getSirketToplams(0, getColName(sheet, c, row2.getRowNum())), style2);
			createFormualCell(row3, c, getSirketToplams(1, getColName(sheet, c, row3.getRowNum())), style2);
			createFormualCell(row4, c, getSirketToplams(2, getColName(sheet, c, row4.getRowNum())), style2);
			createFormualCell(row5, c, "+" + getColIdentifier(sheet, c, row5.getRowNum()) + "/" + getColIdentifier(sheet, c, row4.getRowNum()) + "-1",
					style2perc);
			createFormualCell(row6, c, "+" + getColIdentifier(sheet, c, row5.getRowNum()) + "/" + getColIdentifier(sheet, c, row3.getRowNum()) + "-1",
					style2perc);
			createFormualCell(row7, c, getColName(sheet, c, row7.getRowNum()) + musteriRow, style2simple);

		}
		
		regions.add(new Region(pos-3,(short)(1), pos-2, (short) (c)));

		pos += 2;

		HSSFRow row1a = sheet.createRow((short) pos++);
		HSSFRow row2a = sheet.createRow((short) pos++);
		HSSFRow row3a = sheet.createRow((short) pos++);
		HSSFRow row4a = sheet.createRow((short) pos++);
		HSSFRow row5a = sheet.createRow((short) pos++);
		HSSFRow row6a = sheet.createRow((short) pos++);
		HSSFRow row7a = sheet.createRow((short) pos++);

		short c1 = 0;

		createCell(row1a, c1, "SIRKET TOPLAM", style1);
		createCell(row2a, c1, "NET SATIŞ - GEÇEN YIL", style1left);
		createCell(row3a, c1, "NET SATIŞ - BÜTÇE", style1left);
		createCell(row4a, c1, "NET SATIS -FİİLİ", style1left);
		createCell(row5a, c1, "Prog.Göre Artış", style1left);
		createCell(row6a, c1, "G.Yıla Göre Artış", style1left);
		createCell(row7a, c1, "MUSTERI SAYISI", style1left);

		c1++;

		createCell(row1a, c1, "KUMULE", style1);
		createFormualCell(row2a, c1, "SUM(" + getColIdentifier(sheet, c1, row2.getRowNum() + 1) + ":" + getColIdentifier(sheet, c, row2.getRowNum() + 1) + ")",
				style2);
		createFormualCell(row3a, c1, "SUM(" + getColIdentifier(sheet, c1, row3.getRowNum() + 1) + ":" + getColIdentifier(sheet, c, row3.getRowNum() + 1) + ")",
				style2);
		createFormualCell(row4a, c1, "SUM(" + getColIdentifier(sheet, c1, row4.getRowNum() + 1) + ":" + getColIdentifier(sheet, c, row4.getRowNum() + 1) + ")",
				style2);
		createFormualCell(row5a, c1, "+" + getColIdentifier(sheet, c1, row5a.getRowNum()) + "/" + getColIdentifier(sheet, c1, row4a.getRowNum()) + "-1",
				style2perc);
		createFormualCell(row6a, c1, "+" + getColIdentifier(sheet, c1, row5a.getRowNum()) + "/" + getColIdentifier(sheet, c1, row3a.getRowNum()) + "-1",
				style2perc);
		createFormualCell(row7a, c1, "SUM(" + getColIdentifier(sheet, c1, row7.getRowNum() + 1) + ":" + getColIdentifier(sheet, c, row7.getRowNum() + 1) + ")",
				style2simple);
		
		
		regions.add(new Region(pos-3,(short)(1), pos-2, (short) (1)));

		return pos;
	}

	String[] sirketToplamColums = { "", "", "", "" };

	private String getSirketToplams(int i, String col) {
		String[] rows = sirketToplamColums[i].split(",");
		return col + rows[0] + "+" + col + rows[1];
	}

	private short drawMagazaToplams(HSSFSheet sheet, ResultSet rset) throws SQLException {

		pos += 2;

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row1 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row2 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row3 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row4 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row5 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row6 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row7 = sheet.createRow((short) pos++);

		short c = 0;

		createCell(row1, c, "MAGAZA KASA SATIS TOPLAM", style1);
		createCell(row2, c, "NET SATIŞ - GEÇEN YIL", style1left);
		createCell(row3, c, "NET SATIŞ - BÜTÇE", style1left);
		createCell(row4, c, "NET SATIS -FİİLİ", style1left);
		createCell(row5, c, "Prog.Göre Artış", style1left);
		createCell(row6, c, "G.Yıla Göre Artış", style1left);
		createCell(row7, c, "MUSTERI SAYISI", style1left);

		sirketToplamColums[0] += row2.getRowNum() + 1 + ",";
		sirketToplamColums[1] += row3.getRowNum() + 1 + ",";
		sirketToplamColums[2] += row4.getRowNum() + 1 + ",";
		sirketToplamColums[3] += row7.getRowNum() + 1 + ",";

		while (rset.next()) {

			c++;

			createCell(row1, c, rset.getString(1), style1);
			createCell(row2, c, rset.getDouble(2), style2);
			createCell(row3, c, rset.getDouble(3), style2);
			createCell(row4, c, rset.getDouble(4), style2);
			createFormualCell(row5, c, "+" + getColIdentifier(sheet, c, row5.getRowNum()) + "/" + getColIdentifier(sheet, c, row4.getRowNum()) + "-1",
					style2perc);
			createFormualCell(row6, c, "+" + getColIdentifier(sheet, c, row5.getRowNum()) + "/" + getColIdentifier(sheet, c, row3.getRowNum()) + "-1",
					style2perc);
			createCell(row7, c, rset.getDouble(5), style2simple);

		}
		
		regions.add(new Region(pos-3,(short)(1), pos-2, (short) c));

		return pos;
	}

	private short drawToptanSatis(HSSFSheet sheet, ResultSet rset) throws SQLException {

		pos += 2;

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row1 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row2 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row3 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row4 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row5 = sheet.createRow((short) pos++);

		// addMergedRegion(sheet, wb, pos, 1, pos, 2);
		HSSFRow row6 = sheet.createRow((short) pos++);

		short c = 0;

		createCell(row1, c, "TOPLU SATIS", style1);
		createCell(row2, c, "NET SATIŞ - GEÇEN YIL", style1left);
		createCell(row3, c, "NET SATIŞ - BÜTÇE", style1left);
		createCell(row4, c, "NET SATIS -FİİLİ", style1left);
		createCell(row5, c, "Prog.Göre Artış", style1left);
		createCell(row6, c, "G.Yıla Göre Artış", style1left);

		sirketToplamColums[0] += row2.getRowNum() + 1 + ",";
		sirketToplamColums[1] += row3.getRowNum() + 1 + ",";
		sirketToplamColums[2] += row4.getRowNum() + 1 + ",";

		while (rset.next()) {

			c++;

			createCell(row1, c, rset.getString(1), style1);
			createCell(row2, c, rset.getDouble(2), style2);
			createCell(row3, c, rset.getDouble(3), style2);
			createCell(row4, c, rset.getDouble(4), style2);
			createFormualCell(row5, c, "+" + getColIdentifier(sheet, c, row5.getRowNum()) + "/" + getColIdentifier(sheet, c, row4.getRowNum()) + "-1",
					style2perc);
			createFormualCell(row6, c, "+" + getColIdentifier(sheet, c, row5.getRowNum()) + "/" + getColIdentifier(sheet, c, row3.getRowNum()) + "-1",
					style2perc);

		}
		
		regions.add(new Region(pos-2,(short)(1), pos-1, (short) c));

		return pos;
	}

	private void createFirstSheet() throws SQLException, IOException, WritingNotSupportedException {

		sheet1 = wb.createSheet("Yerel");
		sheet1.createFreezePane(1, 4);

		initStyle();

		HSSFRow header = sheet1.createRow((short) 2);
		header.setHeight((short) 350);

		HSSFRow subheader = sheet1.createRow((short) 3);
		subheader.setHeight((short) 700);

		HSSFRow title = sheet1.createRow((short) 0);
		createCell(title, 0, "RAMSTORE KAZAKISTAN", style1);

		addMergedRegion(sheet1, wb, 2, 0, 3, 0);
		createCell(header, 0, "TARIH", style1);
		setColumnWidth(sheet1, 0, 5000);

		// -------- Drawing in width magazas

		megaTotalColumns[0] = "";
		megaTotalColumns[1] = "";
		megaTotalColumns[2] = "";
		megaTotalColumns[3] = "";

		sirketToplamColums[0] = "";
		sirketToplamColums[1] = "";
		sirketToplamColums[2] = "";
		sirketToplamColums[3] = "";
		
		regions = new ArrayList<Region>();

		int index = 0;

		for (String store : storeList) {

			CallableStatement cstmt4 = conn.prepareCall("call RAM_REPORTS.gsyrel(?,?,?)");
			cstmt4.registerOutParameter(1, OracleTypes.CURSOR);

			cstmt4.setString(2, store);
			cstmt4.setDate(3, new Date(donem.getTime()));
			cstmt4.execute();

			pos = drawMagazaDetails(store, index++, sheet1, header, subheader, (ResultSet) cstmt4.getObject(1));
			

			cstmt4.close();
		}
		
		

		// -------- Drawing in magazas toplam, pivot magaza 100

		{
			CallableStatement cstmt4t = conn.prepareCall("call RAM_REPORTS.gsyrel(?,?,?)");
			cstmt4t.registerOutParameter(1, OracleTypes.CURSOR);

			cstmt4t.setString(2, "100");
			cstmt4t.setDate(3, new Date(donem.getTime()));
			cstmt4t.execute();

			pos = drawMagazaDetailsToplams(index++, sheet1, header, subheader, (ResultSet) cstmt4t.getObject(1));

			cstmt4t.close();
		}

		// -------- Drawing magaza toplam excluding toptan satis
		{

			CallableStatement cstmt5 = conn.prepareCall("call RAM_REPORTS.gsmtyrel(?,?)");
			cstmt5.registerOutParameter(1, OracleTypes.CURSOR);

			cstmt5.setDate(2, new Date(donem.getTime()));
			cstmt5.execute();

			pos = drawMagazaToplams(sheet1, (ResultSet) cstmt5.getObject(1));

			cstmt5.close();

		}

		// -------- Drawing magaza toplam only toptan satis
		{

			CallableStatement cstmt6 = conn.prepareCall("call RAM_REPORTS.gstsyrel(?,?)");
			cstmt6.registerOutParameter(1, OracleTypes.CURSOR);

			cstmt6.setDate(2, new Date(donem.getTime()));
			cstmt6.execute();

			pos = drawToptanSatis(sheet1, (ResultSet) cstmt6.getObject(1));

			cstmt6.close();

		}

		// -------- Drawing megatoplams
		{

			CallableStatement cstmt7 = conn.prepareCall("call RAM_REPORTS.gstsyrel(?,?)");
			cstmt7.registerOutParameter(1, OracleTypes.CURSOR);

			cstmt7.setDate(2, new Date(donem.getTime()));
			cstmt7.execute();

			pos = drawSirketToplams(sheet1, (ResultSet) cstmt7.getObject(1));

			cstmt7.close();

		}
		
		HSSFSheetConditionalFormatting cf = sheet1.getSheetConditionalFormatting();

		HSSFConditionalFormattingRule rule = cf.createConditionalFormattingRule(ComparisonOperator.LT, "0",null);
		HSSFFontFormatting fntFrm = rule.createFontFormatting();
		fntFrm.setFontColorIndex(HSSFFont.COLOR_RED);
		
		Region[] regs = new Region[regions.size()];
		regions.toArray(regs);
		
		cf.addConditionalFormatting(regs, rule);

	}

	private short drawMagazaDetails(String store, int index, HSSFSheet sheet, HSSFRow header, HSSFRow subheader, ResultSet rset) throws SQLException {

		int offset = index * 6;

		short i = 0;
		short r = 4;
		short rAy = 5;
		String tmpAy = "#";
		String[] totals = { "", "", "", "" };
		
		while (rset.next()) {

			if (r == 4) {
				addMergedRegion(sheet, wb, 2, 1 + offset, 2, 6 + offset);
				createCell(header, 1 + offset, rset.getString(4), style1);

				createCell(subheader, 1 + offset, "NET SATIŞ - GEÇEN YIL", style1);
				setColumnWidth(sheet, 1 + offset, 3500);
				createCell(subheader, 2 + offset, "NET SATIŞ - BÜTÇE", style1);
				setColumnWidth(sheet, 2 + offset, 3500);
				createCell(subheader, 3 + offset, "NET SATIS -FİİLİ", style1);
				setColumnWidth(sheet, 3 + offset, 3500);
				createCell(subheader, 4 + offset, "Prog.Göre Artış", style1);
				setColumnWidth(sheet, 4 + offset, 3000);
				createCell(subheader, 5 + offset, "G.Yıla Göre Artış", style1);
				setColumnWidth(sheet, 5 + offset, 3000);
				createCell(subheader, 6 + offset, "MUSTERI SAYISI", style1);
				setColumnWidth(sheet, 6 + offset, 3000);
			}

			HSSFRow row = sheet.createRow(r++);
			
			

			if (!rset.getString(2).equals(tmpAy) && !tmpAy.equals("#")) {

				createCell(row, 0, tmpAy, style1);

				r--;

				createFormualCell(row, 1 + offset, "SUM(" + getColIdentifier(sheet, 1 + offset, rAy) + ":" + getColIdentifier(sheet, 1 + offset, r) + ")",
						style4);
				createFormualCell(row, 2 + offset, "SUM(" + getColIdentifier(sheet, 2 + offset, rAy) + ":" + getColIdentifier(sheet, 2 + offset, r) + ")",
						style4);
				createFormualCell(row, 3 + offset, "SUM(" + getColIdentifier(sheet, 3 + offset, rAy) + ":" + getColIdentifier(sheet, 3 + offset, r) + ")",
						style4);
				createFormualCell(row, 4 + offset, "+" + getColIdentifier(sheet, 3 + offset, (r + 1)) + "/" + getColIdentifier(sheet, 2 + offset, (r + 1))
						+ "-1", style4perc);
				createFormualCell(row, 5 + offset, "+" + getColIdentifier(sheet, 3 + offset, (r + 1)) + "/" + getColIdentifier(sheet, 1 + offset, (r + 1))
						+ "-1", style4perc);
				createFormualCell(row, 6 + offset, "SUM(" + getColIdentifier(sheet, 6 + offset, rAy) + ":" + getColIdentifier(sheet, 6 + offset, r) + ")",
						style4simple);

				// sheet.groupRow(rAy - 1, r - 1);
				// sheet.setRowGroupCollapsed(r-1,true);

				r++;

				totals[0] += getColIdentifier(sheet, 1 + offset, r) + "+";
				totals[1] += getColIdentifier(sheet, 2 + offset, r) + "+";
				totals[2] += getColIdentifier(sheet, 3 + offset, r) + "+";
				totals[3] += getColIdentifier(sheet, 6 + offset, r) + "+";

				row = sheet.createRow(r++);
				rAy = r;

			}

			if (index == 0)
				createCell(row, 0, rset.getString(1), style2centered);

			createCell(row, 1 + offset, rset.getDouble(5), style2);
			createCell(row, 2 + offset, rset.getDouble(6), style2);
			createCell(row, 3 + offset, rset.getDouble(7), style2);
			createFormualCell(row, 4 + offset, "+" + getColIdentifier(sheet, 3 + offset, r) + "/" + getColIdentifier(sheet, 2 + offset, r) + "-1", style2perc);
			createFormualCell(row, 5 + offset, "+" + getColIdentifier(sheet, 3 + offset, r) + "/" + getColIdentifier(sheet, 1 + offset, r) + "-1", style2perc);
			createCell(row, 6 + offset, rset.getDouble(8), style2simple);

			tmpAy = rset.getString(2);

		}

		HSSFRow row = sheet.createRow(r++);

		createCell(row, 0, tmpAy, style1);

		r--;

		createFormualCell(row, 1 + offset, "SUM(" + getColIdentifier(sheet, 1 + offset, rAy) + ":" + getColIdentifier(sheet, 1 + offset, r) + ")", style4);
		createFormualCell(row, 2 + offset, "SUM(" + getColIdentifier(sheet, 2 + offset, rAy) + ":" + getColIdentifier(sheet, 2 + offset, r) + ")", style4);
		createFormualCell(row, 3 + offset, "SUM(" + getColIdentifier(sheet, 3 + offset, rAy) + ":" + getColIdentifier(sheet, 3 + offset, r) + ")", style4);
		createFormualCell(row, 4 + offset, "+" + getColIdentifier(sheet, 3 + offset, (r + 1)) + "/" + getColIdentifier(sheet, 2 + offset, (r + 1)) + "-1",
				style4perc);
		createFormualCell(row, 5 + offset, "+" + getColIdentifier(sheet, 3 + offset, (r + 1)) + "/" + getColIdentifier(sheet, 1 + offset, (r + 1)) + "-1",
				style4perc);
		createFormualCell(row, 6 + offset, "SUM(" + getColIdentifier(sheet, 6 + offset, rAy) + ":" + getColIdentifier(sheet, 6 + offset, r) + ")", style4simple);

		r++;

		totals[0] += getColIdentifier(sheet, 1 + offset, r);
		totals[1] += getColIdentifier(sheet, 2 + offset, r);
		totals[2] += getColIdentifier(sheet, 3 + offset, r);
		totals[3] += getColIdentifier(sheet, 6 + offset, r);

		r++;

		row = sheet.createRow(r++);

		createCell(row, 0, "TOPLAM", style1);

		createFormualCell(row, 1 + offset, totals[0], style4);
		createFormualCell(row, 2 + offset, totals[1], style4);
		createFormualCell(row, 3 + offset, totals[2], style4);
		createFormualCell(row, 4 + offset, getColIdentifier(sheet, 3 + offset, r) + "/" + getColIdentifier(sheet, 2 + offset, r) + "-1", style4perc);
		createFormualCell(row, 5 + offset, getColIdentifier(sheet, 3 + offset, r) + "/" + getColIdentifier(sheet, 1 + offset, r) + "-1", style4perc);
		createFormualCell(row, 6 + offset, totals[3], style4simple);

		megaTotalColumns[0] += getColName(sheet, 1 + offset, r) + ",";
		megaTotalColumns[1] += getColName(sheet, 2 + offset, r) + ",";
		megaTotalColumns[2] += getColName(sheet, 3 + offset, r) + ",";
		megaTotalColumns[3] += getColName(sheet, 6 + offset, r) + ",";

		/*
		 * sheet.autoSizeColumn((short)(1+offset));
		 * sheet.autoSizeColumn((short)(2+offset));
		 * 
		 * sheet.autoSizeColumn((short)(3+offset));
		 * sheet.autoSizeColumn((short)(4+offset));
		 * 
		 * sheet.autoSizeColumn((short)(5+offset));
		 * sheet.autoSizeColumn((short)(6+offset));
		 */
		
		regions.add(new Region(5,(short)(4 + offset), r, (short) (5+offset)));

		return r;

	}

	String[] megaTotalColumns = { null, null, null, null };

	private short drawMagazaDetailsToplams(int index, HSSFSheet sheet, HSSFRow header, HSSFRow subheader, ResultSet rset) throws SQLException {

		int offset = index * 6;

		addMergedRegion(sheet, wb, 2, 1 + offset, 2, 6 + offset);
		createCell(header, 1 + offset, "SIRKET TOPLAM", style1);

		createCell(subheader, 1 + offset, "NET SATIŞ - GEÇEN YIL", style1);
		setColumnWidth(sheet, 1 + offset, 3500);
		createCell(subheader, 2 + offset, "NET SATIŞ - BÜTÇE", style1);
		setColumnWidth(sheet, 2 + offset, 3500);
		createCell(subheader, 3 + offset, "NET SATIS -FİİLİ", style1);
		setColumnWidth(sheet, 3 + offset, 3000);
		createCell(subheader, 4 + offset, "Prog.Göre Artış", style1);
		setColumnWidth(sheet, 4 + offset, 3000);
		createCell(subheader, 5 + offset, "G.Yıla Göre Artış", style1);
		setColumnWidth(sheet, 5 + offset, 3000);
		createCell(subheader, 6 + offset, "MUSTERI SAYISI", style1);
		setColumnWidth(sheet, 6 + offset, 3000);

		short r = 4;
		short rAy = 5;
		String tmpAy = "#";
		String[] totals = { "", "", "", "" };

		while (rset.next()) {
			HSSFRow row = sheet.createRow(r++);

			if (!rset.getString(2).equals(tmpAy) && !tmpAy.equals("#")) {

				createCell(row, 0, tmpAy, style1);

				r--;

				createFormualCell(row, 1 + offset, "SUM(" + getColIdentifier(sheet, 1 + offset, rAy) + ":" + getColIdentifier(sheet, 1 + offset, r) + ")",
						style4);
				createFormualCell(row, 2 + offset, "SUM(" + getColIdentifier(sheet, 2 + offset, rAy) + ":" + getColIdentifier(sheet, 2 + offset, r) + ")",
						style4);
				createFormualCell(row, 3 + offset, "SUM(" + getColIdentifier(sheet, 3 + offset, rAy) + ":" + getColIdentifier(sheet, 3 + offset, r) + ")",
						style4);
				createFormualCell(row, 4 + offset, "+" + getColIdentifier(sheet, 3 + offset, (r + 1)) + "/" + getColIdentifier(sheet, 2 + offset, (r + 1))
						+ "-1", style4perc);
				createFormualCell(row, 5 + offset, "+" + getColIdentifier(sheet, 3 + offset, (r + 1)) + "/" + getColIdentifier(sheet, 1 + offset, (r + 1))
						+ "-1", style4perc);
				createFormualCell(row, 6 + offset, "SUM(" + getColIdentifier(sheet, 6 + offset, rAy) + ":" + getColIdentifier(sheet, 6 + offset, r) + ")",
						style4simple);

				sheet.groupRow(rAy - 1, r - 1);
				sheet.setRowGroupCollapsed(r - 1, true);

				r++;

				totals[0] += getColIdentifier(sheet, 1 + offset, r) + "+";
				totals[1] += getColIdentifier(sheet, 2 + offset, r) + "+";
				totals[2] += getColIdentifier(sheet, 3 + offset, r) + "+";
				totals[3] += getColIdentifier(sheet, 6 + offset, r) + "+";

				row = sheet.createRow(r++);
				rAy = r;

			}

			createCell(row, 0, rset.getString(1), style2centered);

			// System.out.println(getMegaTotals(3,r));

			createFormualCell(row, 1 + offset, getMegaTotals(0, r), style2);
			createFormualCell(row, 2 + offset, getMegaTotals(1, r), style2);
			createFormualCell(row, 3 + offset, getMegaTotals(2, r), style2);
			createFormualCell(row, 4 + offset, "+" + getColIdentifier(sheet, 3 + offset, r) + "/" + getColIdentifier(sheet, 2 + offset, r) + "-1", style2perc);
			createFormualCell(row, 5 + offset, "+" + getColIdentifier(sheet, 3 + offset, r) + "/" + getColIdentifier(sheet, 1 + offset, r) + "-1", style2perc);
			createFormualCell(row, 6 + offset, getMegaTotals(3, r), style2simple);

			tmpAy = rset.getString(2);

		}

		HSSFRow row = sheet.createRow(r++);

		createCell(row, 0, tmpAy, style1);

		r--;

		createFormualCell(row, 1 + offset, "SUM(" + getColIdentifier(sheet, 1 + offset, rAy) + ":" + getColIdentifier(sheet, 1 + offset, r) + ")", style4);
		createFormualCell(row, 2 + offset, "SUM(" + getColIdentifier(sheet, 2 + offset, rAy) + ":" + getColIdentifier(sheet, 2 + offset, r) + ")", style4);
		createFormualCell(row, 3 + offset, "SUM(" + getColIdentifier(sheet, 3 + offset, rAy) + ":" + getColIdentifier(sheet, 3 + offset, r) + ")", style4);
		createFormualCell(row, 4 + offset, "+" + getColIdentifier(sheet, 3 + offset, (r + 1)) + "/" + getColIdentifier(sheet, 2 + offset, (r + 1)) + "-1",
				style4perc);
		createFormualCell(row, 5 + offset, "+" + getColIdentifier(sheet, 3 + offset, (r + 1)) + "/" + getColIdentifier(sheet, 1 + offset, (r + 1)) + "-1",
				style4perc);
		createFormualCell(row, 6 + offset, "SUM(" + getColIdentifier(sheet, 6 + offset, rAy) + ":" + getColIdentifier(sheet, 6 + offset, r) + ")", style4simple);

		r++;

		totals[0] += getColIdentifier(sheet, 1 + offset, r);
		totals[1] += getColIdentifier(sheet, 2 + offset, r);
		totals[2] += getColIdentifier(sheet, 3 + offset, r);
		totals[3] += getColIdentifier(sheet, 6 + offset, r);

		r++;

		row = sheet.createRow(r++);

		createCell(row, 0, "TOPLAM", style1);

		createFormualCell(row, 1 + offset, totals[0], style4);
		createFormualCell(row, 2 + offset, totals[1], style4);
		createFormualCell(row, 3 + offset, totals[2], style4);
		createFormualCell(row, 4 + offset, getColIdentifier(sheet, 3 + offset, r) + "/" + getColIdentifier(sheet, 2 + offset, r) + "-1", style4perc);
		createFormualCell(row, 5 + offset, getColIdentifier(sheet, 3 + offset, r) + "/" + getColIdentifier(sheet, 1 + offset, r) + "-1", style4perc);
		createFormualCell(row, 6 + offset, totals[3], style4simple);

		/*
		 * sheet.autoSizeColumn((short)(1+offset));
		 * sheet.autoSizeColumn((short)(2+offset));
		 * 
		 * sheet.autoSizeColumn((short)(3+offset));
		 * sheet.autoSizeColumn((short)(4+offset));
		 * 
		 * sheet.autoSizeColumn((short)(5+offset));
		 * sheet.autoSizeColumn((short)(6+offset));
		 */
		
		regions.add(new Region(5,(short)(4 + offset), r, (short) (5+offset)));

		return r;

	}

	private String getMegaTotals(int col, short row) {
		String formula = "";
		String[] cols = megaTotalColumns[col].split(",");
		for (String colName : cols) {
			formula += "+" + colName + "" + row;
		}
		return formula;
	}

	public String getColName(HSSFSheet sheet, int col, int row) {
		String id = new CellReference(row - 1, col, false, false).formatAsString();
		return id.substring(0, id.length() - String.valueOf(row).length());
	}

	public String getColIdentifier(HSSFSheet sheet, int col, int row) {
		// System.out.println(new
		// CellReference(row,col,false,false).formatAsString());
		return new CellReference(row - 1, col, false, false).formatAsString();
		// return columnNames[col]+row;
	}

	private void createFormualCell(HSSFRow row, int col, String formula, HSSFCellStyle style) {
		HSSFCell cell = row.createCell((short) col);
		cell.setCellFormula(formula);
		cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);
		cell.setCellStyle(style);
	}

	// Special negative with red
	private void createFormualCell(HSSFRow row, int col, String formula, HSSFCellStyle style, Double value) {
		HSSFCell cell = row.createCell((short) col);
		cell.setCellFormula(formula);
		cell.setCellType(HSSFCell.CELL_TYPE_FORMULA);

		if (value < 0) {
			HSSFFont font = wb.createFont();
			font.setFontHeightInPoints((short) 8);
			font.setFontName("Arial");
			font.setColor(HSSFFont.COLOR_RED);
			style.setFont(font);
			cell.setCellStyle(style);
		} else {
			cell.setCellStyle(style);
		}

	}

	private void setColumnWidth(HSSFSheet sheet, int column, int width) {
		sheet.setColumnWidth((short) column, (short) width);
	}

	private void createCell(HSSFRow row, int col, String value, HSSFCellStyle style) {
		HSSFCell cell = row.createCell((short) col);
		cell.setCellType(HSSFCell.CELL_TYPE_STRING);
		cell.setCellStyle(style);
		cell.setCellValue(value);
	}

	private void createCell(HSSFRow row, int col, Double value, HSSFCellStyle style) {
		HSSFCell cell = row.createCell((short) col);
		cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
		cell.setCellStyle(style);
		cell.setCellValue(value);
	}

	private void addMergedRegion(HSSFSheet sheet, HSSFWorkbook wb, int a, int b, int c, int d) {
		Region region = new Region(a, (short) b, c, (short) d);
		sheet.addMergedRegion(region);
		/*
		 * try {
		 * HSSFRegionUtil.setBorderBottom(HSSFCellStyle.BORDER_THIN,region,sheet,wb); }
		 * catch (Exception e) { }
		 */
	}

	private void initStyle() {
		// Create a new font and alter it.
		HSSFFont font1 = wb.createFont();
		font1.setFontHeightInPoints((short) 8);
		font1.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font1.setFontName("Arial");

		HSSFFont font2 = wb.createFont();
		font2.setFontHeightInPoints((short) 8);
		font2.setFontName("Arial");

		HSSFFont font3 = wb.createFont();
		font3.setFontHeightInPoints((short) 10);
		font3.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font3.setFontName("Arial");

		HSSFDataFormat format = wb.createDataFormat();
		short fmt1 = format.getFormat("#,##0.00");
		// short fmt2 = format.getFormat("#,##0");
		short fmt3 = format.getFormat("### ### ### ### ##0");
		short fmt4 = format.getFormat("### ### ### ### ##0");
		short fmt5 = format.getFormat("0.00%");

		style1 = wb.createCellStyle();
		style1.setFillForegroundColor(HSSFColor.GOLD.index);
		style1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style1.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style1.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style1.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style1.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style1.setFont(font1);
		style1.setWrapText(true);

		style2 = wb.createCellStyle();
		style2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style2.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style2.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		style2.setFont(font2);
		style2.setDataFormat(fmt3);

		style2simple = wb.createCellStyle();
		style2simple.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style2simple.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style2simple.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style2simple.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style2simple.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		style2simple.setFont(font2);
		style2simple.setDataFormat(fmt4);

		style2centered = wb.createCellStyle();
		style2centered.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style2centered.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style2centered.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style2centered.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style2centered.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		style2centered.setFont(font2);
		style2centered.setDataFormat(fmt1);

		style2perc = wb.createCellStyle();
		style2perc.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style2perc.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style2perc.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style2perc.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style2perc.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		style2perc.setFont(font2);
		style2perc.setDataFormat(fmt5);

		style4 = wb.createCellStyle();
		style4.setFillForegroundColor(HSSFColor.GOLD.index);
		style4.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style4.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style4.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style4.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style4.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style4.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		style4.setFont(font1);
		style4.setDataFormat(fmt3);

		style4perc = wb.createCellStyle();
		style4perc.setFillForegroundColor(HSSFColor.GOLD.index);
		style4perc.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style4perc.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style4perc.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style4perc.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style4perc.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style4perc.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		style4perc.setFont(font1);
		style4perc.setDataFormat(fmt5);

		style4simple = wb.createCellStyle();
		style4simple.setFillForegroundColor(HSSFColor.GOLD.index);
		style4simple.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style4simple.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style4simple.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style4simple.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style4simple.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style4simple.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		style4simple.setFont(font1);
		style4simple.setDataFormat(fmt4);

		style1left = wb.createCellStyle();
		style1left.setFillForegroundColor(HSSFColor.GOLD.index);
		style1left.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style1left.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style1left.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style1left.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style1left.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style1left.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		style1left.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style1left.setFont(font1);
		style1left.setWrapText(true);

		style8Text = wb.createCellStyle();
		// style8Text.setAlignment(HSSFCellStyle.ALIGN_RIGHT );
		style8Text.setFont(font1);
		style8Text.setDataFormat(fmt1);

	}
	
	@Override
	public String getFileFormat() {
		return "dd.MM.yyyy";
	}

}
