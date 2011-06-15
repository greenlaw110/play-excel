package controllers;

import play.mvc.Before;
import play.mvc.Controller;

/**
 * ExcelFormat extends PlayFramework Sessoin.format according to http accept header
 * 
 * @author luog
 *
 */
public class ExcelFormat extends Controller {
    @Before
    public static void extendFormat() {
        if (request.headers.get("accept") != null) {
            String accept = request.headers.get("accept").value();
            if (accept.indexOf("text/csv") != -1) request.format = "csv";
            if (accept.matches(".*application\\/(excel|vnd\\.ms\\-excel|x\\-excel|x\\-msexcel).*")) request.format = "xls";
            if (accept.indexOf("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") != -1) request.format = "xlsx";
        }
    }
}
