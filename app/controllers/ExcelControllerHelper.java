package controllers;

import java.io.UnsupportedEncodingException;

import org.apache.commons.codec.net.URLCodec;

import play.PlayPlugin;
import play.jobs.Job;
import play.jobs.OnApplicationStart;
import play.libs.F;
import play.libs.F.Promise;
import play.modules.excel.Plugin;
import play.modules.excel.RenderExcel;
import play.mvc.After;
import play.mvc.Before;
import play.mvc.Controller;
import play.mvc.Scope;
import play.mvc.Scope.RenderArgs;
import play.templates.Template;
import play.vfs.VirtualFile;

/**
 * ExcelFormat extends PlayFramework Sessoin.format according to http accept
 * header
 * 
 * @author luog
 * 
 */
public class ExcelControllerHelper extends Controller {
    @Before
    public static void extendFormat() {
        if (request.headers.get("accept") != null) {
            String accept = request.headers.get("accept").value();
            if (accept.indexOf("text/csv") != -1)
                request.format = "csv";
            if (accept
                    .matches(".*application\\/(excel|vnd\\.ms\\-excel|x\\-excel|x\\-msexcel).*"))
                request.format = "xls";
            if (accept
                    .indexOf("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") != -1)
                request.format = "xlsx";
        }
    }

    private static final URLCodec encoder = new URLCodec();

    @After
    public static void setHeader() throws UnsupportedEncodingException {
        if (null != request.format && !request.format.matches("(csv|xls|xlsx)"))
            return;

        if (!response.headers.containsKey("Content-Disposition")) {
            String fileName = renderArgs.get(RenderExcel.RA_FILENAME,
                    String.class);
            if (fileName == null) {
                response.setHeader("Content-Disposition",
                        "attachment; filename=export." + request.format);
            } else {
                response.setHeader(
                        "Content-Disposition",
                        "attachment; filename="
                                + encoder.encode(fileName, "utf-8"));
            }

            if (request.format.equals("xls")) {
                response.setContentTypeIfNotSet("application/vnd.ms-excel");
            } else if (request.format.equals("xlsx")) {
                response.setContentTypeIfNotSet("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            } else if (request.format.equals("csv")) {
                response.setContentTypeIfNotSet("text/csv");
            }
        }
    }

    // constants copied from ActionInvoker
    public static Template loadTemplate(VirtualFile file) {
        if (RenderExcel.async()) {
            Promise<RenderExcel> render = RenderExcel.renderAsync(file.relativePath(), Scope.RenderArgs.current().data, null);
            await(render, new F.Action<RenderExcel>() {
                @Override
                public void invoke(RenderExcel result) {
                    RenderArgs renderArgs = RenderArgs.current();
                    if (!renderArgs.data.containsKey(RenderExcel.RA_FILENAME)) {
                        renderArgs.put(RenderExcel.RA_FILENAME, result.getFileName());
                    }
                    throw result;
                }
            });
            return null;
        } else {
            return new Plugin.ExcelTemplate(file);
        }
    }

    @OnApplicationStart
    public static class BootLoader extends Job<Object> {
        @Override
        public void doJob() {
            play.modules.excel.Plugin.templateLoader = new PlayPlugin() {
                public play.templates.Template loadTemplate(VirtualFile file) {
                    return ExcelControllerHelper.loadTemplate(file);
                }
            };
        }
    }
}
