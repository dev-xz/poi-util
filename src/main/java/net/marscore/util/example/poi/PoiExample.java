package net.marscore.util.example.poi;

import net.marscore.util.poi.ExcelUtil;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Eden
 */
public class PoiExample {
    public static void main(String[] args) {
        System.out.println("--读取学生.xlsx信息并输出--");
        // 获取resources下的文件
        String studentFilePath = PoiExample.class.getResource("/example/Student.xlsx").getFile();
        List<Student> students = (List<Student>)ExcelUtil.excelFileToList(studentFilePath, Student.class);
        for(Student student:students){
            System.out.println(student);
        }
        System.out.println("--读取教师.xls信息并输出--");
        String teacherFilePath = PoiExample.class.getResource("/example/Teacher.xls").getFile();
        List<Teacher> teachers = (List<Teacher>)ExcelUtil.excelFileToList(teacherFilePath, Teacher.class);
        for(Teacher teacher:teachers){
            System.out.println(teacher);
        }
        // 导出在当前目录运行盘的根目录下
        ExcelUtil.listToFile("/TeacherExport.xlsx", teachers);
        // 仅导出指定的字段，并赋予别名
        Map<String, String> myField = new HashMap<String, String>();
        myField.put("account", "学号");
        myField.put("name", "姓名");
        ExcelUtil.listToFile("/StudentExport", students, myField);
    }
}