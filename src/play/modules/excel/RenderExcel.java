package play.modules.excel;

import java.io.InputStream;
import java.util.Map;

import net.sf.jxls.transformer.XLSTransformer;

import org.apache.commons.codec.net.URLCodec;
import org.apache.poi.ss.usermodel.Workbook;

import play.Play;
import play.exceptions.UnexpectedException;
import play.mvc.Http.Request;
import play.mvc.Http.Response;
import play.mvc.results.Result;
import play.vfs.VirtualFile;

/**
 * 200 OK with application/excel
 * 
 * This Result try to render Excel file with given template and beans map The
 * code use jxls and poi library to render Excel
 */
@SuppressWarnings("serial")
public class RenderExcel extends Result {

    private static URLCodec encoder = new URLCodec();
    private static VirtualFile tmplRoot = null;
    String templateName = null;
    String fileName = null; // recommended report file name
    boolean inline = true;
    Map<String, Object> beans = null;

    private static void initTmplRoot() {
        VirtualFile appRoot = VirtualFile.open(Play.applicationPath);
        String rootDef = "";
        if (Play.configuration.containsKey("excel.template.root"))
            rootDef = (String) Play.configuration.get("excel.template.root");
        tmplRoot = appRoot.child(rootDef);
    }

    public RenderExcel(String templateName, Map<String, Object> beans) {
        this(templateName, beans, null);
    }

    public RenderExcel(String templateName, Map<String, Object> beans,
            String fileName) {
        this(templateName, beans, fileName, false);
    }

    public RenderExcel(String templateName, Map<String, Object> beans,
            String fileName, boolean inline) {
        this.templateName = templateName;
        this.inline = inline;
        this.beans = beans;
        this.fileName = fileName == null ? fileName_(templateName) : fileName;
    }
    
    private static String fileName_(String path) {
        int i = path.lastIndexOf("/");
        if (-1 == i) return path;
        return path.substring(++i);
    }
    
    public static void main(String[] args) {
        System.out.println(fileName_("abc.xls"));
        System.out.println(fileName_("/xyz/abc.xls"));
        System.out.println(fileName_("app/xyz/abc.xls"));
    }

    @Override
    public void apply(Request request, Response response) {
        try {
            if (!response.headers.containsKey("Content-Disposition")) {
                if (inline) {
                    if (fileName == null) {
                        response.setHeader("Content-Disposition", "inline");
                    } else {
                        response.setHeader(
                                "Content-Disposition",
                                "inline; filename="
                                        + encoder.encode(fileName, "utf-8"));
                    }
                } else if (fileName == null) {
                    response.setHeader("Content-Disposition",
                            "attachment; filename=export.xls");
                } else {
                    response.setHeader(
                            "Content-Disposition",
                            "attachment; filename="
                                    + encoder.encode(fileName, "utf-8"));
                }
            }
            if (templateName.matches(".*\\.(xlsx|xlxlsm|xltx|xltm)")) {
                setContentTypeIfNotSet(response,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            } else {
                setContentTypeIfNotSet(response, "application/vnd.ms-excel");
            }

            if (null == tmplRoot) {
                initTmplRoot();
            }
            InputStream is = tmplRoot.child(templateName).inputstream();
            Workbook workbook = new XLSTransformer().transformXLS(is, beans);
            workbook.write(response.out);
            is.close();
        } catch (Exception e) {
            throw new UnexpectedException(e);
        }
    }
}
