package net.marscore.util.example.poi;

import java.util.Objects;

/**
 * @author Eden
 */
public class Student {
    private String account;
    private String name;
    private String stuClass;

    public String getAccount() {
        return account;
    }

    public void setAccount(String account) {
        this.account = account;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getStuClass() {
        return stuClass;
    }

    public void setStuClass(String stuClass) {
        this.stuClass = stuClass;
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        Student student = (Student) o;
        return Objects.equals(account, student.account) &&
                Objects.equals(name, student.name) &&
                Objects.equals(stuClass, student.stuClass);
    }

    @Override
    public int hashCode() {
        return Objects.hash(account, name, stuClass);
    }

    @Override
    public String toString() {
        return "Student{" +
                "account='" + account + '\'' +
                ", name='" + name + '\'' +
                ", stuClass='" + stuClass + '\'' +
                '}';
    }
}
