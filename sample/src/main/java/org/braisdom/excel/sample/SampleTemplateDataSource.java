package org.braisdom.excel.sample;

import org.braisdom.excel.TemplateDataSource;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Set;

public class SampleTemplateDataSource implements TemplateDataSource {

    public Set<String> getDataNames() {
        return Collections.singleton("users");
    }

    public Object getData(String name) {
        List<User> users = new ArrayList<User>();
        User abbas = new User();
        abbas.setName("Abbas");
        abbas.setGender("male");
        abbas.setOccupation("Software Engineer");
        abbas.setAge(11);

        User almeric = new User();
        almeric.setName("Almeric");
        almeric.setGender("male");
        almeric.setOccupation("Software Engineer");
        almeric.setAge(12);

        users.add(abbas);
        users.add(almeric);

        return users;
    }

    public static class User {
        private String name;
        private String gender;
        private String occupation;
        private Integer age;

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getGender() {
            return gender;
        }

        public void setGender(String gender) {
            this.gender = gender;
        }

        public String getOccupation() {
            return occupation;
        }

        public void setOccupation(String occupation) {
            this.occupation = occupation;
        }

        public Integer getAge() {
            return age;
        }

        public void setAge(Integer age) {
            this.age = age;
        }
    }

}
