package controllers;

import play.PlayPlugin;
import play.jobs.Job;
import play.jobs.OnApplicationStart;
import play.libs.F;
import play.libs.F.Promise;
import play.modules.excel.Plugin;
import play.modules.excel.RenderExcel;
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
