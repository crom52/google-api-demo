package com.demo.vessel;

import com.google.api.services.gmail.Gmail;
import com.google.api.services.gmail.Gmail.Users.Messages;
import com.google.api.services.gmail.model.Message;
import com.google.api.services.gmail.model.MessagePart;
import static com.google.common.base.Strings.emptyToNull;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import static java.time.ZonedDateTime.now;
import static java.time.format.DateTimeFormatter.ofPattern;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.StringJoiner;
import static java.util.function.Function.identity;
import static java.util.stream.Collectors.toMap;
import java.util.stream.IntStream;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import static org.apache.commons.lang3.StringUtils.isBlank;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;
import static org.springframework.util.ObjectUtils.nullSafeToString;

@Component
@RequiredArgsConstructor
@Slf4j
class VesselScheduleService {
  @Qualifier("gmailService")
  private final Gmail service;
  @Value("${gmailApi.sender:phuc52crom@gmail.com}")
  private String sender;
  @Value("${gmailApi.subject:Test Google Email}")
  private String subject;

  List<?> findNext7DaysVesselPlan() {
    try {
      Messages messageService = service.users().messages();
      List<Message> emails = getTodayEmails(messageService, subject, sender);
      String user = "me";
      List<Object> resultList = new ArrayList<>();
      for (var email : emails) {
        MessagePart attachment = getAttachmentExcel(messageService, email, user);
        Sheet sheet = getSheetInExelFile(service, email.getId(), attachment.getBody().getAttachmentId());
        List<String> columnNames = getColumnNames(sheet.getRow(0));
        var etaAfter7DayRows = getRowsHasETAAfter7Days(sheet, columnNames);
        resultList.add(convertToMaps(columnNames, etaAfter7DayRows));
      }
      return resultList;
    } catch (Exception e) {
      return List.of();
    }
  }

  private List<Message> getTodayEmails(Messages messageService, String subject, String sender) throws IOException {
    if (messageService == null || emptyToNull(subject) == null || emptyToNull(sender) == null) {
      return List.of();
    }
    LocalDate today = LocalDate.now();
    var startDate = today.atStartOfDay(ZoneId.systemDefault()).toInstant();
    var endDate = today.atStartOfDay(ZoneId.systemDefault()).plusDays(1).toInstant().minusMillis(1);
    String user = "me";
    String query = buildQueryEmail(subject, sender, startDate, endDate);
    return messageService.list(user).setQ(query).execute().getMessages();
  }

  private MessagePart getAttachmentExcel(Messages messageService, Message email, String user) throws IOException {
    if (messageService == null || email == null || isBlank(user)) {
      throw new IllegalArgumentException("Failed getAttachmentExcel");
    }
    var content = messageService.get(user, email.getId()).execute();
    return content
        .getPayload()
        .getParts()
        .stream()
        .filter(att -> att.getFilename().startsWith("vs"))
        .findFirst()
        .orElseThrow();
  }

  private Sheet getSheetInExelFile(Gmail service, String messageId, String attachmentId)
      throws IOException, InvalidFormatException {
    if (service == null || isBlank(messageId) || isBlank(attachmentId)) {
      throw new IllegalArgumentException("Failed getSheetInExelFile");
    }

    byte[] data = service.users().messages().attachments().get("me", messageId, attachmentId).execute().decodeData();
    Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(data));
    return workbook.getSheetAt(0);
  }

  private List<String> getColumnNames(Row headerRow) {
    if (headerRow == null) {
      throw new IllegalArgumentException("Failed getColumnNames");
    }
    return IntStream
        .range(0, headerRow.getLastCellNum())
        .mapToObj(headerRow::getCell)
        .map(cell -> cell != null ? cell.getStringCellValue() : "column_")
        .toList();
  }

  private List<Row> getRowsHasETAAfter7Days(Sheet sheet, List<String> columnNames) {
    if (columnNames.isEmpty() || sheet == null) {
      throw new IllegalArgumentException("Failed getRowsHasETAAfter7Days");
    }
    var indexOfEtaColumn = columnNames.indexOf("ETA");
    return IntStream
        .range(1, sheet.getLastRowNum() + 1)
        .mapToObj(sheet::getRow)
        .filter(Objects::nonNull)
        .filter(e -> isAfter7Days(e.getCell(indexOfEtaColumn).toString()))
        .toList();
  }

  private List<Map<String, String>> convertToMaps(List<String> columnNames, List<Row> etaAfter7DayRows) {
    if (columnNames.isEmpty() || etaAfter7DayRows.isEmpty()) {
      return List.of();
    }

    return etaAfter7DayRows
        .stream()
        .map(row -> columnNames
            .stream()
            .collect(toMap(identity(), key -> nullSafeToString(row.getCell(columnNames.indexOf(key))))))
        .toList();
  }

  private String buildQueryEmail(String subject, String sender, Instant startDate, Instant endDate) {
    if (isBlank(subject) || isBlank(sender) || startDate == null || endDate == null) {
      throw new IllegalArgumentException("Failed buildQueryEmail");
    }
    StringJoiner query = new StringJoiner(" ");
    query
        .add("subject:" + subject)
        .add("from:" + sender)
        .add("before:" + endDate.getEpochSecond())
        .add("after:" + startDate.getEpochSecond());
    return query.toString();
  }

  private boolean isAfter7Days(String etaStr) {
    if (isBlank(etaStr)) {
      return false;
    }
    var etaDateTime = ZonedDateTime.parse(etaStr, ofPattern("yyyyMMdd HHmm").withZone(ZoneId.systemDefault()));
    return etaDateTime.isAfter(now().withZoneSameInstant(ZoneId.systemDefault()).plusDays(7));
  }

}
