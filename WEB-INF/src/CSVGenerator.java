import java.io.IOException;
import java.net.URLDecoder;
import javax.servlet.http.*;

import com.dhtmlx.xml2excel.CSVWriter;

public class CSVGenerator extends HttpServlet {

		public void doPost(HttpServletRequest req, HttpServletResponse resp) throws IOException {
			String xml = req.getParameter("grid_xml");
			xml = URLDecoder.decode(xml, "UTF-8");
			CSVWriter writer = new CSVWriter();
			writer.generate(xml, resp);
		}
}
