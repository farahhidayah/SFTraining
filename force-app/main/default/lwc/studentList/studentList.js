import { LightningElement, wire } from 'lwc';
import { loadScript } from 'lightning/platformResourceLoader';
import SHEETJS from '@salesforce/resourceUrl/sheetjs';

import getContacts from '@salesforce/apex/ContactController.getContacts';

export default class StudentList extends LightningElement {
    sheetjsInitialized = false;

    // Data
    contacts = [];
    students = [];

    // Checkbox state
    studentExport = false;
    contactExport = false;

    // Columns for display (UI only, not affecting export)
    studentColumns = [
        { label: 'Name', fieldName: 'Name' },
        { label: 'Age', fieldName: 'Age', type: 'number' },
        { label: 'Course', fieldName: 'Course' }
    ];

    contactColumns = [
        { label: 'First Name', fieldName: 'Name' },
        { label: 'Email', fieldName: 'Email' }
    ];

    // Wire Apex for contacts
    @wire(getContacts)
    wiredContacts({ error, data }) {
        if (data) {
            this.contacts = data;
            console.log('Contacts fetched', data);
        } else if (error) {
            console.error('Error fetching contacts', error);
        }
    }

    // Hardcode students
    connectedCallback() {
        this.students = [
            { Name: 'John Doe', Age: 22, Course: 'Computer Science' },
            { Name: 'Jane Smith', Age: 24, Course: 'Mathematics' },
            { Name: 'Ali Ahmad', Age: 21, Course: 'Engineering' }
        ];
    }

    // Load SheetJS
    renderedCallback() {
        if (this.sheetjsInitialized) {
            return;
        }
        this.sheetjsInitialized = true;

        loadScript(this, SHEETJS + '/xlsx.full.min.js')
            .then(() => {
                console.log('✅ SheetJS loaded successfully', typeof XLSX);
            })
            .catch(error => {
                console.error('❌ Error loading SheetJS', error);
            });
    }

    // Checkbox handlers
    handleStudentCheckbox(event) {
        this.studentExport = event.target.checked;
    }

    handleContactCheckbox(event) {
        this.contactExport = event.target.checked;
    }

    // Export logic
    exportToExcel() {
        if (typeof XLSX === 'undefined') {
            console.error('❌ XLSX not loaded yet');
            return;
        }

        const workbook = XLSX.utils.book_new();

        // ✅ Default: export both if no box checked
        if (!this.studentExport && !this.contactExport) {
            this.addStudentSheet(workbook);
            this.addContactSheet(workbook);
        } else {
            if (this.studentExport) {
                this.addStudentSheet(workbook);
            }
            if (this.contactExport) {
                this.addContactSheet(workbook);
            }
        }

        XLSX.writeFile(workbook, 'Data.xlsx');
    }

    // Helpers to add sheets with totals
    addStudentSheet(workbook) {
        const data = [...this.students];
        data.push({ Name: 'Total Students', Age: data.length, Course: '' });
        const sheet = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(workbook, sheet, 'Students');
    }

    addContactSheet(workbook) {
        const data = [...this.contacts];
        data.push({ Id: 'Total Contacts', Name: data.length, Email: '' });
        const sheet = XLSX.utils.json_to_sheet(data);
        XLSX.utils.book_append_sheet(workbook, sheet, 'Contacts');
    }
}
