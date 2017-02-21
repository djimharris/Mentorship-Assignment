//Read excel files and match mentors to mentees


if(typeof require !== 'undefined') XLSX = require('xlsx');
var write_excel = require('./write_excel.js')

function Mentors(id, name, email, dept, spec){
	this.id = id;
	this.name = name;
	this.email = email;
	this.dept = dept;
	this.spec = spec;
	this.mentees = [];
}

function Mentees(id, name, email, dept, spec){
	this.id = id;
	this.name = name;
	this.email = email;
	this.dept = dept;
	this.spec = spec;
	this.mentor = "";
}

var menteeLimit = 3;
var freeMentors = [];
var listMentors = [];

var mentorsBook = XLSX.readFile('mentors.xlsx');	//Read excel files containing mentors
var mentorsSheetName = mentorsBook.SheetNames[0];
var row = 1;
var column = 'A';
var address_of_cell = 'A3';
var mentorsSheet = mentorsBook.Sheets[mentorsSheetName];

var col_mentor_name = "";
var col_mentor_email = "";
var col_mentor_dept = "";
var col_mentor_spec = "";

while(mentorsSheet[column + row] != null){			//Assign column numbers to the required values
	value = mentorsSheet[column + row].v;
	if(value.toLowerCase().trim() === "name"){
		col_mentor_name = column;
	}
	if(value.toLowerCase().trim() === "email"){
		col_mentor_email = column;
	}
	if(value.toLowerCase() === "department"){
		col_mentor_dept = column;
	}
	if(value.toLowerCase() === "specialisation"){
		col_mentor_spec = column;
	}
	column = String.fromCharCode(column.charCodeAt(0) + 1);
}
row = 2;
id = 0;

while(mentorsSheet[col_mentor_name + row] != null){
	name = mentorsSheet[col_mentor_name + row].v;
	email = mentorsSheet[col_mentor_email + row].v;
	dept = mentorsSheet[col_mentor_dept + row].v;
	spec = mentorsSheet[col_mentor_spec + row].v;
	temp_mentor = new Mentors(id, name, email, dept, spec);
	listMentors.push(temp_mentor);
	freeMentors.push(temp_mentor);
	row++;
	id++;
}

var listMentees = [];
var freeMentees = [];

var menteesBook = XLSX.readFile('mentees.xlsx');	//Read excel files containing mentees
var menteesSheetName = menteesBook.SheetNames[0];
var row = 1;
var column = 'A';
var address_of_cell = 'A3';
var menteesSheet = menteesBook.Sheets[menteesSheetName];

var col_mentee_name = "";
var col_mentee_email = "";
var col_mentee_dept = "";
var col_mentee_spec = "";

while(menteesSheet[column + row] != null){			//Assign column numbers to the required values
	value = menteesSheet[column + row].v;
	if(value.toLowerCase().trim() === "name"){
		col_mentee_name = column;
	}
	if(value.toLowerCase().trim() === "email"){
		col_mentee_email = column;
	}
	if(value.toLowerCase() === "department"){
		col_mentee_dept = column;
	}
	if(value.toLowerCase() === "specialisation"){
		col_mentee_spec = column;
	}
	column = String.fromCharCode(column.charCodeAt(0) + 1);
}
row = 2;
id = 0;

while(menteesSheet[col_mentee_name + row] != null){
	name = menteesSheet[col_mentee_name + row].v;
	email = menteesSheet[col_mentee_email + row].v;
	dept = menteesSheet[col_mentee_dept + row].v;
	spec = menteesSheet[col_mentee_spec + row].v;
	temp_mentee = new Mentees(id, name, email, dept, spec)
	listMentees.push(temp_mentee);
	freeMentees.push(temp_mentee);
	row++;
	id++;
}


for (x of listMentees){
	if(x.mentor === ""){
		validMentors = [];
		for(y of freeMentors){
			if((x.spec === y.spec) && (x.dept === y.dept)){
				validMentors.push(y);
			}
		}
		if(validMentors.length > 0){
			// randomAssignment = Math.floor(Math.random() * validMentors.length);
			validMentors.sort(function(mentor1, mentor2){return mentor1.mentees.length - mentor2.mentees.length})
			randomAssignment = 0;
			x.mentor = validMentors[randomAssignment];
			validMentors[randomAssignment].mentees.push(x);
			freeMentees.splice(freeMentees.indexOf(x), 1);
			if(validMentors[randomAssignment].mentees.length == menteeLimit){
				freeMentors.splice(freeMentors.indexOf(validMentors[randomAssignment]), 1)
			}
		}
	}

}

for (x of listMentees){
	if(x.mentor === ""){
		validMentors = [];
		for(y of freeMentors){
			if(x.dept === y.dept){
				validMentors.push(y);
			}
		}
		if(validMentors.length > 0){
			// randomAssignment = Math.floor(Math.random() * validMentors.length);
			validMentors.sort(function(mentor1, mentor2){return mentor1.mentees.length - mentor2.mentees.length})
			randomAssignment = 0;
			x.mentor = validMentors[randomAssignment];
			validMentors[randomAssignment].mentees.push(x);
			freeMentees.splice(freeMentees.indexOf(x), 1);
			if(validMentors[randomAssignment].mentees.length == menteeLimit){
				freeMentors.splice(freeMentors.indexOf(validMentors[randomAssignment]), 1)
			}
		}
	}

}

for (x of listMentees){
	if(x.mentor === ""){
		if(freeMentors.length > 0){
			// randomAssignment = Math.floor(Math.random() * freeMentors.length);
			freeMentors.sort(function(mentor1, mentor2){return mentor1.mentees.length - mentor2.mentees.length})
			randomAssignment = 0;
			x.mentor = freeMentors[randomAssignment];
			freeMentors[randomAssignment].mentees.push(x);
			freeMentees.splice(freeMentees.indexOf(x), 1);
			if(freeMentors[randomAssignment].mentees.length == menteeLimit){
				freeMentors.splice(freeMentors.indexOf(freeMentors[randomAssignment]), 1)
			}
		}
	}

}


// for(x of listMentees){
// 	console.log(x.name + " is assigned to " + x.mentor.name);
// }

var mentorData = [["S.no", "Assigned Mentees", "Name", "Email", "Department", "Specialisation"]];
for(x of listMentors){
	var string_of_mentees = "";
	for(y of x.mentees){
		string_of_mentees = string_of_mentees + y.id + " ";
	}
	mentorData.push([x.id, string_of_mentees , x.name, x.email, x.dept, x.spec]);
}
var mentorsSheetName = "Mentors";

var menteeData = [["S.no", "Assigned Mentees", "Name", "Email", "Department", "Specialisation"]];
for(x of listMentees){
	menteeData.push([x.id, x.mentor.id , x.name, x.email, x.dept, x.spec]);
}

var menteesSheetName = "Mentees";

write_excel.custom_write = new custom_write(mentorData, mentorsSheetName, menteeData, menteesSheetName);

