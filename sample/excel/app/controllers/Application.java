package controllers;

import static play.modules.excel.Excel.renderExcel;

import java.util.List;

import models.Contact;
import play.mvc.Controller;

public class Application extends Controller {

    public static void index() {
    	List<Contact> contacts = Contact.findAll();
        render(contacts);
    }
    
    public static void generateNameCard(Long id) {
    	Contact person = Contact.findById(id);
    	renderArgs.put("fileName", person.getEntityId() + ".xls");
    	renderExcel(person);
    }

}              
