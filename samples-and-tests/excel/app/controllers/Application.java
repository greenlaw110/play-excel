package controllers;

import java.util.Date;
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
    	String fileName = person.getId() + ".xls";
    	request.format = "xls";
    	render(fileName, person);
    }
    
    public static void generateAddressBook() {
    	List<Contact> contacts = Contact.findAll();
    	Date date = new Date();
    	String fileName = "address_book.xls";
        render(fileName, contacts, date);
    }

}              
