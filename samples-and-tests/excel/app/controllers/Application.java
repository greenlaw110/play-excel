package controllers;

import java.util.Date;
import java.util.List;

import models.Contact;
import play.Logger;
import play.modules.excel.RenderExcel;
import play.mvc.Controller;
import play.mvc.With;

@With(ExcelControllerHelper.class)
public class Application extends Controller {

    public static void index() {
    	List<Contact> contacts = Contact.findAll();
        render(contacts);
    }
    
    public static void generateNameCard(Long id) {
        Logger.info("generateNameCard");
    	Contact person = Contact.findById(id);
    	request.format = "xls";
        renderArgs.put("__EXCEL_FILE_NAME__", person.getId() + ".xls");
    	render(person);
    }
    
    public static void generateAddressBook() {
        Logger.info("generateAddressBook");
    	List<Contact> contacts = Contact.findAll();
    	Date date = new Date();
    	String __EXCEL_FILE_NAME__ = "address_book.xlsx";
    	renderArgs.put(RenderExcel.RA_ASYNC, true);
        render(__EXCEL_FILE_NAME__, contacts, date);
        Logger.info("generateAddressBook exit");
    }

}
