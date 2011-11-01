package com.dhtmlx.xml2excel;

import java.io.IOException;
import java.io.PrintWriter;

import javax.servlet.http.HttpServletResponse;

public class HTMLWriter extends BaseWriter{
	int rows = 0;
	int cols = 0;
	int fontSize = -1;
	String watermark = null;

	public void generate(String xml, HttpServletResponse resp) throws IOException {
		CSVxml data = new CSVxml(xml);
		
		resp.setCharacterEncoding("UTF-8");
		resp.setContentType("application/vnd.ms-excel");
		resp.setCharacterEncoding("UTF-8");
		resp.setHeader("Content-Disposition", "attachment;filename=grid.xls");
		resp.setHeader("Cache-Control", "max-age=0");
	
		String[] csv;
		int colsnum = 0;
		PrintWriter writer = resp.getWriter();

		startHTML(writer);
		csv = data.getHeader();
		if (csv != null)colsnum = csv.length;
		while(csv != null){			
			writer.append(dataAsString(csv));
			csv = data.getHeader();
		}
		
		csv = data.getRow();
		if (csv != null){
			colsnum = csv.length;
			cols = csv.length;
		}
		while(csv != null){			
			writer.append(dataAsString(csv));
			csv = data.getRow();
			rows +=1;
		}
		
		csv = data.getFooter();
		if (csv != null)colsnum = csv.length;
		while(csv != null){			
			writer.append(dataAsString(csv));
			writer.flush();
			csv = data.getFooter();
		}
		drawWatermark(writer, colsnum);
		endHTML(writer);
		
		writer.flush();
		writer.close();
	}

	private void drawWatermark(PrintWriter writer, int colsnum) {
		if (watermark != null)
			writer.append("<tr><td colspan='" + colsnum + "'>" + watermark + "</td></tr>");
	}
	
	private void endHTML(PrintWriter writer) {
		writer.append("</table></body></html>");
	}

	private void startHTML(PrintWriter writer) {
		writer.append("<html><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><body>");
		if (fontSize != -1)
			writer.append("<style>table tr td { font-size: " + fontSize + "px; }</style>");
		writer.append("<table>");
		
	}

	private String dataAsString(String[] csv) {
		if (csv.length == 0) return "";
		
		StringBuffer buff = new StringBuffer();
		buff.append("<tr>");
		for ( int i=0; i<csv.length; i++){
			buff.append("<td>");
			buff.append(csv[i].replace("&", "&amp;").replace(">", "&gt;").replace("<", "&lt;"));
			buff.append("</td>");				
		}	
		buff.append("</tr>\n");
		return buff.toString();
	}

	public int getColsStat() {
		return cols;
	}

	public int getRowsStat() {
		return rows;
	}
	
	public void setFontSize(int fontsize) {
		this.fontSize = fontsize;
	}

	public void setWatermark(String watermark) {
		this.watermark = watermark;
	}
}
