package models;

import javax.persistence.Entity;

import play.db.jpa.Model;

@Entity
public class Contact extends Model {

	private static final long serialVersionUID = -5806820381004668188L;

	public String firstName;
	public String lastName;
	public String title;
	public String address;
	public String mobile;
	public String email;
	
	public String toString() {
		return title + " " + firstName + " " + lastName; 
	}
}
