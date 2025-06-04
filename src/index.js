const xlsx = window.XLSX;

function xlsToDate(serial) {
  const excelEpoch = new Date(1899, 11, 30);
  return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
}

function xlsTime(excelTime, ref) {
  const date = ref || new Date();
  const hours = Math.floor(excelTime * 24);
  const minutes = Math.floor((excelTime * 24 * 60) % 60);
  const seconds = Math.floor((excelTime * 24 * 60 * 60) % 60);

  return new Date(
    date.getFullYear(),
    date.getMonth(),
    date.getDate(),
    hours,
    minutes,
    seconds
  );
}

function timeToXls(date) {
  const hours = date.getHours();
  const minutes = date.getMinutes();
  const seconds = date.getSeconds();
  const milliseconds = date.getMilliseconds();
  return (hours * 3600 + minutes * 60 + seconds + milliseconds / 1000) / 86400;
}

function isAvailable(staff, examStart, examEnd) {
  return staff.schedule.every((slot) => {
    return examEnd <= slot.startTime || examStart >= slot.endTime;
  });
}

function getDurationInHours(start, end) {
  return (end - start) / (1000 * 60 * 60);
}

const MAX_WORKLOAD = 8;

function assignStaffToExams(exams, staffs) {
  const assignments = [];

  const workload = staffs.map((staff) => ({
    ...staff,
    assignedHours: 0,
  }));

  for (const exam of exams) {
    const examDuration = getDurationInHours(exam.startTime, exam.endTime);

    const assignedStaff = []; // Sort staff by least workload

    const sortedStaff = workload
      .filter(
        (staff) =>
          isAvailable(staff, exam.startTime, exam.endTime) &&
          staff.assignedHours + examDuration <= MAX_WORKLOAD
      )
      .sort((a, b) => a.assignedHours - b.assignedHours);

    for (let i = 0; i < exam.workload && i < sortedStaff.length; i++) {
      const staff = sortedStaff[i];
      staff.assignedHours += examDuration;
      staff.schedule.push({ startTime: exam.startTime, endTime: exam.endTime });
      assignedStaff.push(staff.name);
    }

    assignments.push({
      ...exam,
      assignedStaff,
    });
  }

  return assignments;
}

function downloadAssignments(assignments) {
  const worksheet = xlsx.utils.json_to_sheet(
    assignments.map((assignment) => ({
      Exam: assignment.name,
      Date: assignment.startTime,
      "Start Time": timeToXls(assignment.startTime),
      "End Time": timeToXls(assignment.endTime),
      Staffs: assignment.workload,
      "Assigned Staffs": assignment.assignedStaff.join(", "),
    }))
  );

  const range = xlsx.utils.decode_range(worksheet["!ref"]);
  for (let R = range.s.r + 1; R <= range.e.r; ++R) {
    const cellS = xlsx.utils.encode_cell({ r: R, c: 2 });
    const cellE = xlsx.utils.encode_cell({ r: R, c: 3 });
    for (const cellAddress of [cellS, cellE]) {
      if (worksheet[cellAddress]) {
        worksheet[cellAddress].z = "hh:mm";
      }
    }
  }

  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "Assignments");
  xlsx.writeFile(workbook, "exam_assignment.xlsx");
}

const input = document.getElementById("spreadsheet");
input.addEventListener("change", () => {
  if (input.files[0] == null) return;
  const reader = new FileReader();
  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = xlsx.read(data, { type: "array" });

    const exams = [];
    const staffs = [];
    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const json = xlsx.utils.sheet_to_json(worksheet);
      if (sheetName === "Exams") {
        for (const row of json) {
          const date = xlsToDate(row["Date"]);
          exams.push({
            name: row["Name"],
            workload: row["Workload"],
            startTime: xlsTime(row["Start Time"], date),
            endTime: xlsTime(row["End Time"], date),
          });
        }
      } else if (sheetName.startsWith("Staff - ")) {
        const name = sheetName.replace(/^Staff - /, "");
        const schedule = [];
        for (const row of json) {
          const date = xlsToDate(row["Date"]);
          schedule.push({
            startTime: xlsTime(row["Start Time"], date),
            endTime: xlsTime(row["End Time"], date),
          });
        }
        staffs.push({ name, schedule });
      }
    }

    const assignments = assignStaffToExams(exams, staffs);
    downloadAssignments(assignments);
  };
  reader.readAsArrayBuffer(input.files[0]);
});
