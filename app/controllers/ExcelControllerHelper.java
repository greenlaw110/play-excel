package controllers;

import play.PlayPlugin;
import play.classloading.enhancers.LocalvariablesNamesEnhancer;
import play.jobs.Job;
import play.jobs.OnApplicationStart;
import play.libs.F;
import play.libs.F.Promise;
import play.modules.excel.Plugin;
import play.modules.excel.RenderExcel;
import play.mvc.Controller;
import play.mvc.Scope;
import play.mvc.Scope.RenderArgs;
import play.mvc.Util;
import play.templates.Template;
import play.vfs.VirtualFile;

import java.util.List;

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

    public static void renderDynamicFull(List objects, String beanName, int sheetStart, List sheetNames, String sheetNamePrefix, String sheetNameSuffix, Object... args) {
        renderArgs.put(RenderExcel.RA_DYNAMIC_OBJECTS, objects);
        renderArgs.put(RenderExcel.RA_DYNAMIC_BEAN_NAME, beanName);
        renderArgs.put(RenderExcel.RA_DYNAMIC_SHEET_START, sheetStart);
        renderArgs.put(RenderExcel.RA_DYNAMIC_SHEET_NAMES, sheetNames);
        renderArgs.put(RenderExcel.RA_DYNAMIC_SHEET_NAME_PREFIX, sheetNamePrefix);
        renderArgs.put(RenderExcel.RA_DYNAMIC_SHEET_NAME_SUFFIX, sheetNameSuffix);
        String templateName = null;
        if (args.length > 6 && args[6] instanceof String && LocalvariablesNamesEnhancer.LocalVariablesNamesTracer.getAllLocalVariableNames(args[6]).isEmpty()) {
            templateName = args[6].toString();
        } else {
            templateName = template();
        }
        renderTemplate(templateName, args);
    }

    @Util
    public static void renderDynamic(List objects, String beanName, int sheetStart, Object... args) {
        renderArgs.put(RenderExcel.RA_DYNAMIC_OBJECTS, objects);
        renderArgs.put(RenderExcel.RA_DYNAMIC_BEAN_NAME, beanName);
        renderArgs.put(RenderExcel.RA_DYNAMIC_SHEET_START, sheetStart);
        String templateName = null;
        if (args.length > 3 && args[3] instanceof String && LocalvariablesNamesEnhancer.LocalVariablesNamesTracer.getAllLocalVariableNames(args[3]).isEmpty()) {
            templateName = args[3].toString();
        } else {
            templateName = template();
        }
        renderTemplate(templateName, args);
    }
}
