/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.mavenproject1;



    public class User  {
        private String name;
        private String email;
        private String school;
        private int age;
        private String gender;

    public User(String name, String email, String school, int age, String gender) {
        this.name = name;
        this.email = email;
        this.school = school;
        this.age = age;
        this.gender = gender;
    }
   
    public User(){
   
   
   
   }

   
   
   
    @Override
    public String toString() {
        return "User{" + "name=" + name + ", email=" + email + ", school=" + school + ", age=" + age + ", gender=" + gender + '}';
    }


    public String getName() {
        return name;
    }

    public String getEmail() {
        return email;
    }


    public String getSchool() {
        return school;
    }

    public int getAge() {
        return age;
    }

    public String getGender() {
        return gender;
    }
    

    public void setName(String name) {
        this.name = name;
    }

    public void setEmail(String email) {
        this.email = email;
    }

    public void setSchool(String school) {
        this.school = school;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    
    
}
