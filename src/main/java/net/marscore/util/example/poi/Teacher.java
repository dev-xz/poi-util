package net.marscore.util.example.poi;

import java.util.Objects;

/**
 * @author Eden
 */
public class Teacher {
    private String account;
    private String name;

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

    @Override
    public String toString() {
        return "Teacher{" +
                "account='" + account + '\'' +
                ", name='" + name + '\'' +
                '}';
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) {
            return true;
        }
        if (o == null || getClass() != o.getClass()) {
            return false;
        }
        Teacher teacher = (Teacher) o;
        return Objects.equals(account, teacher.account) &&
                Objects.equals(name, teacher.name);
    }

    @Override
    public int hashCode() {

        return Objects.hash(account, name);
    }
}
