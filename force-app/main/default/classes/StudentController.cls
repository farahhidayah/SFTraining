public with sharing class StudentController {
    @AuraEnabled(cacheable=true)
    public static List<Student__c> getStudents(){
        return [SELECT Name, Email__c, Course__c FROM Student__c LIMIT 100];
    }
}
