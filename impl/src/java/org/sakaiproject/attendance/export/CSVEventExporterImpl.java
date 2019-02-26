/*
 *  Copyright (c) 2017, University of Dayton
 *
 *  Licensed under the Educational Community License, Version 2.0 (the "License");
 *  you may not use this file except in compliance with the License.
 *  You may obtain a copy of the License at
 *
 *              http://opensource.org/licenses/ecl2
 *
 *  Unless required by applicable law or agreed to in writing, software
 *  distributed under the License is distributed on an "AS IS" BASIS,
 *  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *  See the License for the specific language governing permissions and
 *  limitations under the License.
 */

package org.sakaiproject.attendance.export;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.CellType;
import lombok.extern.slf4j.Slf4j;
import lombok.Setter;
import org.sakaiproject.attendance.export.util.SortNameUserComparator;
import org.sakaiproject.attendance.logic.AttendanceLogic;
import org.sakaiproject.attendance.logic.SakaiProxy;
import org.sakaiproject.attendance.model.AttendanceEvent;
import org.sakaiproject.attendance.model.AttendanceStatus;
import org.sakaiproject.attendance.model.Status;
import org.sakaiproject.user.api.User;

import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Collections;
import java.util.Date;
import java.util.List;

/**
 * Implementation of CSVEventExporter, {@link org.sakaiproject.attendance.export.CSVEventExporter}
 *
 * @author Miguel Pellicer [mpellicer (at) edf (dot) global]
 */
@Slf4j
public class CSVEventExporterImpl implements CSVEventExporter {

    private AttendanceEvent event;
    private List<User> users;
    private String groupOrSiteTitle;

    /**
     * {@inheritDoc}
     */
    public void createSignInCsv(AttendanceEvent event, OutputStream outputStream, List<User> usersToPrint, String groupOrSiteTitle) {

        this.event = event;
        this.users = usersToPrint;
        this.groupOrSiteTitle = groupOrSiteTitle;

        buildDocumentShell(outputStream, true);
    }

    /**
     * {@inheritDoc}
     */
    public void createAttendanceSheetCsv(AttendanceEvent event, OutputStream outputStream, List<User> usersToPrint, String groupOrSiteTitle) {

        this.event = event;
        this.users = usersToPrint;
        this.groupOrSiteTitle = groupOrSiteTitle;

        buildDocumentShell(outputStream, false);
    }

    private void buildDocumentShell(OutputStream outputStream, boolean isSignInSheet) {
        String eventName = event.getName();
        Date eventDate = event.getStartDateTime();

        SimpleDateFormat dateFormat = new SimpleDateFormat("EEE, MMM d, yyyy h:mm a");

        try {
            String pageTitle = isSignInSheet?"Sign-In Sheet":"Attendance Sheet";
            String eventDateString = eventDate==null?"":" (" + dateFormat.format(eventDate) + ")";
			int currentRow = 0;
            final HSSFWorkbook wb = new HSSFWorkbook();
            // Create new sheet
            HSSFSheet mainSheet = wb.createSheet("Export");

			// Info rows
			HSSFRow titleRow = sheet.createRow(currentRow);
			HSSFCell titleCell = headerRow.createCell(0);
			titleCell.setCellValue(pageTitle);
			titleCell.setCellType(CellType.STRING);
			currentRow++;

			HSSFRow dateRow = sheet.createRow(currentRow);
			HSSFCell dateCell = dateRow.createCell(0);
			dateCell.setCellValue(eventDateString);
			studentNameCell.setCellType(CellType.STRING);

            if(isSignInSheet) {
                signInSheetTable(mainSheet);
            } else {
                attendanceSheetTable(mainSheet);
            }

            wb.write(outputStream);

            outputStream.close();
            wb.close();

        } catch (final Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void signInSheetTable(HSSFSheet sheet) {

        int currentRow = 2;
        // Create the Header row
        HSSFRow headerRow = sheet.createRow(currentRow);
		HSSFCell studentNameCell = headerRow.createCell(0);
		studentNameCell.setCellValue("Student Name");
		studentNameCell.setCellType(CellType.STRING);
		HSSFCell signatureCell = headerRow.createCell(1);
		signatureCell.setCellValue("Signature");
		signatureCell.setCellType(CellType.STRING);

        Collections.sort(users, new SortNameUserComparator());

        for(User user : users) {
			currentRow++;
			HSSFRow row = sheet.createRow(currentRow);
			studentNameCell = row.createCell(0);
			studentNameCell.setCellValue(user.getSortName());
			studentNameCell.setCellType(CellType.STRING);
			signatureCell = row.createCell(1);
			signatureCell.setCellValue("");
			signatureCell.setCellType(CellType.STRING);
        }
    }

    private void attendanceSheetTable(HSSFSheet sheet) {
		System.out.println("fetido adams");
    }

    /**
     * init - perform any actions required here for when this bean starts up
     */
    public void init() {
        log.debug("CSVEventExporterImpl init()");
    }

    // TODO: Internationalize status header abbreviations
    private String getStatusString(Status s, int numStatuses) {
        if(numStatuses < 4) {
            switch (s)
            {
                case UNKNOWN: return "None";
                case PRESENT: return "Present";
                case EXCUSED_ABSENCE: return "Excused";
                case UNEXCUSED_ABSENCE: return "Absent";
                case LATE: return "Late";
                case LEFT_EARLY: return "Left Early";
                default: return "None";
            }
        } else {
            switch (s)
            {
                case UNKNOWN: return "None";
                case PRESENT: return "Pres";
                case EXCUSED_ABSENCE: return "Excu";
                case UNEXCUSED_ABSENCE: return "Abse";
                case LATE: return "Late";
                case LEFT_EARLY: return "Left";
                default: return "None";
            }
        }
    }

    @Setter
    private SakaiProxy sakaiProxy;

    @Setter
    private AttendanceLogic attendanceLogic;

}
