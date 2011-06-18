package models;

import javax.persistence.Entity;

import play.db.jpa.Model;
import play.jobs.Job;
import play.jobs.OnApplicationStart;
import play.test.Fixtures;

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
        // Contact c = Contact.find("").first();
        // c = Contact.findById(c);
        return title + " " + firstName + " " + lastName;
    }

    @OnApplicationStart
    public static class BootLoader extends Job<Object> {
        @Override
        public void doJob() {
            Fixtures.delete(Contact.class);
            Fixtures.loadModels("initial-data.yml");
        }
    }
}
