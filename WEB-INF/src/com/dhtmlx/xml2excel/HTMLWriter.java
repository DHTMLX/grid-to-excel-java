package com.dhtmlx.xml2excel;

import java.io.IOException;
import java.io.PrintWriter;

import javax.servlet.http.HttpServletResponse;

public class HTMLWriter {
	int rows = 0;
	int cols = 0;
		
	public void generate(String xml, HttpServletResponse resp) throws IOException {
		CSVxml data = new CSVxml(xml);
		
		resp.setCharacterEncoding("UTF-8");
		resp.setContentType("application/vnd.ms-excel");
		resp.setCharacterEncoding("UTF-8");
		resp.setHeader("Content-Disposition", "attachment;filename=grid.xls");
		resp.setHeader("Cache-Control", "max-age=0");
	
		String[] csv;
		PrintWriter writer = resp.getWriter();

		startHTML(writer);
		csv = data.getHeader();
		while(csv != null){			
			writer.append(dataAsString(csv));
			csv = data.getHeader();
		}
		
		csv = data.getRow();
		while(csv != null){			
			writer.append(dataAsString(csv));
			csv = data.getRow();
		}
		
		csv = data.getFooter();
		while(csv != null){			
			writer.append(dataAsString(csv));
			writer.flush();
			csv = data.getFooter();
		}
		endHTML(writer);
		
		writer.flush();
		writer.close();
	}

	private void endHTML(PrintWriter writer) {
		writer.append("</table></body></html>");
		
	}

	private void startHTML(PrintWriter writer) {
		writer.append("<html><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><body><table>");
		
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

}
