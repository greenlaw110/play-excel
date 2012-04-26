package play.modules.excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Map;
import java.util.List;

import net.sf.jxls.transformer.XLSTransformer;

import org.apache.poi.ss.usermodel.Workbook;

import play.Logger;
import play.Play;
import play.exceptions.UnexpectedException;
import play.jobs.Job;
import play.libs.F.Promise;
import play.mvc.Http.Request;
import play.mvc.Http.Response;
import play.mvc.Scope.RenderArgs;
import play.mvc.results.Result;
import play.vfs.VirtualFile;
import play.db.Model;

/**
 * 200 OK with application/excel
 *
 * This Result try to render Excel file with given template and beans map The
 * code use jxls and poi library to render Excel
 */
@SuppressWarnings("serial")
public class RenderExcel extends Result {

    public static final String RA_FILENAME = "__FILE_NAME__";
    public static final String RA_ASYNC = "__EXCEL_ASYNC__";
    public static final String RA_DYNAMIC_OBJECTS = "__EXCEL_DYNAMIC_OBJECTS__";
    public static final String RA_DYNAMIC_SHEET_NAMES = "__EXCEL_DYNAMIC_SHEET_NAMES__";
    public static final String RA_DYNAMIC_SHEET_START = "__EXCEL_DYNAMIC_SHEET_START__";
    public static final String RA_DYNAMIC_SHEET_NAME_PREFIX = "__EXCEL_DYNAMIC_SHEET_NAME_PREFIX__";
    public static final String RA_DYNAMIC_SHEET_NAME_SUFFIX = "__EXCEL_DYNAMIC_SHEET_NAME_SUFFIX__";
    public static final String RA_DYNAMIC_BEAN_NAME = "__EXCEL_DYNAMIC_BEAN_NAME__";
    public static final String CONF_ASYNC = "excel.async";

    private static VirtualFile tmplRoot = null;
    String templateName = null;
    String fileName = null; // recommended report file name
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
        this.templateName = templateName;
        this.beans = beans;
        this.fileName = fileName == null ? fileName_(templateName) : fileName;
    }

    public String getFileName() {
        return fileName;
    }

    public static boolean async() {
        Object o = null;
        if (RenderArgs.current().data.containsKey(RA_ASYNC)) {
            o = RenderArgs.current().get(RA_ASYNC);
        } else {
            o = Play.configuration.get(CONF_ASYNC);
        }
        boolean async = true;
        if (null == o)
            async = true;
        else if (o instanceof Boolean)
            async = (Boolean)o;
        else
            async = Boolean.parseBoolean(o.toString());
        return async;
    }

    private static String fileName_(String path) {
        if (RenderArgs.current().data.containsKey(RA_FILENAME))
            return RenderArgs.current().get(RA_FILENAME, String.class);
        int i = path.lastIndexOf("/");
        if (-1 == i)
            return path;
        return path.substring(++i);
    }

    public static void main(String[] args) {
        System.out.println(fileName_("abc.xls"));
        System.out.println(fileName_("/xyz/abc.xls"));
        System.out.println(fileName_("app/xyz/abc.xls"));
    }

    @Override
    public void apply(Request request, Response response) {
        if (null == excel) {
            Logger.debug("use sync excel rendering");
            long start = System.currentTimeMillis();
            try {
                if (null == tmplRoot) {
                    initTmplRoot();
                }
                InputStream is = tmplRoot.child(templateName).inputstream();
                Workbook workbook = getWorkBook(RenderArgs.current().data);
                workbook.write(response.out);
                is.close();
                Logger.debug("Excel sync render takes %sms", System.currentTimeMillis() - start);
            } catch (Exception e) {
                throw new UnexpectedException(e);
            }
        } else {
            Logger.debug("use async excel rendering...");
            try {
                response.out.write(excel);
            } catch (IOException e) {
                throw new UnexpectedException(e);
            }
        }
    }

    private byte[] excel = null;

    private Workbook getWorkBook(Map<String, Object> renderArgs) throws Exception {
        InputStream is = tmplRoot.child(templateName).inputstream();
        Workbook workbook = null;
        RenderArgs ra = new RenderArgs();
        ra.data = renderArgs;
        if (renderArgs.containsKey(RA_DYNAMIC_OBJECTS)) {
            // dynamic worksheet generation
            // see http://jxls.sourceforge.net/reference/dynamicsheets.html
            List objects = ra.get(RA_DYNAMIC_OBJECTS, List.class);
            List sheetNames = ra.get(RA_DYNAMIC_SHEET_NAMES, List.class);
            Integer sheetStart = ra.get(RA_DYNAMIC_SHEET_START, Integer.class);
            if (null == sheetStart) sheetStart = 0;
            String beanName = ra.get(RA_DYNAMIC_BEAN_NAME, String.class);
            if (null == sheetNames) {
                // try to generate sheetNames
                sheetNames = new ArrayList();
                String prefix = ra.get(RA_DYNAMIC_SHEET_NAME_PREFIX, String.class);
                String suffix = ra.get(RA_DYNAMIC_SHEET_NAME_SUFFIX, String.class);
                int i = 0;
                for (Object o: objects) {
                    StringBuilder name = new StringBuilder();
                    if (o instanceof Model) {
                        Model m = (Model) o;
                        name.append(prefix).append(m._key()).append(suffix);
                    } else {
                        name.append(null == prefix ? "sheet" : prefix).append(i).append(suffix);
                    }
                    sheetNames.add(name.toString());
                    if (null == beanName) beanName = o.getClass().getSimpleName().toLowerCase();
                }
                return new XLSTransformer().transformMultipleSheetsList(is, objects, sheetNames, beanName, renderArgs, sheetStart);
            }
        } else {
            workbook = new XLSTransformer().transformXLS(is, beans);
        }
        return workbook;
    }

    public void preRender() {
        try {
            if (null == tmplRoot) {
                initTmplRoot();
            }
            InputStream is = tmplRoot.child(templateName).inputstream();
            Workbook workbook = getWorkBook(beans);
            ByteArrayOutputStream os = new ByteArrayOutputStream();
            workbook.write(os);
            excel = os.toByteArray();
            is.close();
        } catch (Exception e) {
            throw new UnexpectedException(e);
        }
    }

    public static Promise<RenderExcel> renderAsync(final String templateName, final Map<String, Object> beans, final String fileName) {
        final String fn = fileName == null ? fileName_(templateName) : fileName;
        return new Job<RenderExcel>(){
            @Override
            public RenderExcel doJobWithResult() throws Exception {
                RenderExcel excel = new RenderExcel(templateName, beans, fn);
                excel.preRender();
                return excel;
            }
        }.now();
    }

}
