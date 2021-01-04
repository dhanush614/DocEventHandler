package sample.actionhandler;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.filenet.api.collection.ContentElementList;
import com.filenet.api.collection.FolderSet;
import com.filenet.api.constants.AutoClassify;
import com.filenet.api.constants.AutoUniqueName;
import com.filenet.api.constants.CheckinType;
import com.filenet.api.constants.DefineSecurityParentage;
import com.filenet.api.constants.PropertyNames;
import com.filenet.api.constants.RefreshMode;
import com.filenet.api.core.ContentTransfer;
import com.filenet.api.core.Document;
import com.filenet.api.core.Factory;
import com.filenet.api.core.Folder;
import com.filenet.api.core.ObjectStore;
import com.filenet.api.core.ReferentialContainmentRelationship;
import com.filenet.api.engine.EventActionHandler;
import com.filenet.api.events.ObjectChangeEvent;
import com.filenet.api.property.FilterElement;
import com.filenet.api.property.Properties;
import com.filenet.api.property.PropertyFilter;
import com.filenet.api.util.Id;
import com.ibm.casemgmt.api.Case;
import com.ibm.casemgmt.api.CaseType;
import com.ibm.casemgmt.api.constants.ModificationIntent;
import com.ibm.casemgmt.api.context.CaseMgmtContext;
import com.ibm.casemgmt.api.context.P8ConnectionCache;
import com.ibm.casemgmt.api.context.SimpleP8ConnectionCache;
import com.ibm.casemgmt.api.context.SimpleVWSessionCache;
import com.ibm.casemgmt.api.objectref.ObjectStoreReference;
import com.ibm.casemgmt.api.properties.CaseMgmtProperties;

public class DocumentsEventHandler implements EventActionHandler {
	public void onEvent(ObjectChangeEvent event, Id subId) {
		System.out.println("Inside onEvent method");
		CaseMgmtContext origCmctx = null;
		try {
			P8ConnectionCache connCache = new SimpleP8ConnectionCache();
			origCmctx = CaseMgmtContext.set(new CaseMgmtContext(new SimpleVWSessionCache(), connCache));
			ObjectStore os = event.getObjectStore();
			System.out.println("OS" + os);
			ObjectStoreReference targetOsRef = new ObjectStoreReference(os);
			System.out.println("TOS" + targetOsRef);
			Id id = event.get_SourceObjectId();
			FilterElement fe = new FilterElement(null, null, null, "Owner Name", null);
			PropertyFilter pf = new PropertyFilter();
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_SIZE, null));
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.CONTENT_ELEMENTS, null));
			pf.addIncludeProperty(new FilterElement(null, null, null, PropertyNames.FOLDERS_FILED_IN, null));
			pf.addIncludeProperty(fe);
			Document doc = Factory.Document.fetchInstance(os, id, pf);
			System.out.println("Document Name" + doc.get_Name());
			ContentElementList docContentList = doc.get_ContentElements();
			Iterator iter = docContentList.iterator();
			while (iter.hasNext()) {
				ContentTransfer ct = (ContentTransfer) iter.next();
				InputStream stream = ct.accessContentStream();
				int rowLastCell = 0;
				HashMap<Integer, String> headers = new HashMap<Integer, String>();
				XSSFWorkbook workbook = new XSSFWorkbook(stream);
				XSSFSheet sheet = workbook.getSheetAt(0);
				Iterator<Row> rowIterator = sheet.iterator();
				String headerValue;
				if (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					Iterator<Cell> cellIterator = row.cellIterator();
					int colNum = 0;
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						headerValue = cell.getStringCellValue();
						if (headerValue.contains("*")) {
							if (headerValue.contains("datetime")) {
								headerValue = headerValue.replaceAll("\\* *\\([^)]*\\) *", "").trim();
								headerValue += "dateField";
							} else {
								headerValue = headerValue.replaceAll("\\* *\\([^)]*\\) *", "").trim();
							}
						}
						if (headerValue.contains("datetime")) {
							headerValue = headerValue.replaceAll("\\([^)]*\\) *", "").trim();
							headerValue += "dateField";
						} else {
							headerValue = headerValue.replaceAll("\\([^)]*\\) *", "").trim();
						}
						headers.put(colNum++, headerValue);
					}
					rowLastCell = row.getLastCellNum();
					Cell cell1 = row.createCell(rowLastCell, Cell.CELL_TYPE_STRING);
					if (row.getRowNum() == 0) {
						cell1.setCellValue("Status");
					}

				}
				System.out.println("CaseType Creation");
				CaseType caseType = CaseType.fetchInstance(targetOsRef, doc.get_Name());
				System.out.println("Rights" + caseType.hasInstanceCreationRights());
				while (rowIterator.hasNext()) {
					Row row = rowIterator.next();
					int colNum = 0;
					String caseId = "";
					try {
						Case pendingCase = Case.createPendingInstance(caseType);
						CaseMgmtProperties caseMgmtProperties = pendingCase.getProperties();
						for (int i = 0; i < row.getLastCellNum(); i++) {
							Cell cell = row.getCell(i, Row.CREATE_NULL_AS_BLANK);
							try {
								if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) {
									colNum++;
								} else {
									if (headers.get(colNum).contains("dateField")) {
										System.out.println("Date Field");
										if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC
												&& DateUtil.isCellDateFormatted(cell)) {
											System.out.println("Date Formatted");
											String symName = headers.get(colNum).replace("dateField", "");
											Date date = cell.getDateCellValue();
											System.out.println("Key" + symName + "Value" + date.toString());
											caseMgmtProperties.putObjectValue(symName, date);
											colNum++;
										} else {
											colNum++;
										}
									} else {
										System.out.println("Key1" + headers.get(colNum) + "Value1" + getCharValue(cell));
										caseMgmtProperties.putObjectValue(headers.get(colNum++), getCharValue(cell));
									}
								}
							} catch (Exception e) {
								System.out.println(e);
								e.printStackTrace();
							}
						}
						System.out.println("Case Creation");
						pendingCase.save(RefreshMode.REFRESH, null, null);
						caseId = pendingCase.getId().toString();
						System.out.println("Case_ID: " + caseId);

					} catch (Exception e) {
						System.out.println(e);
						e.printStackTrace();
					}
					Cell cell1 = row.createCell(rowLastCell);
					if (!caseId.isEmpty()) {
						cell1.setCellValue("Success");
					} else {
						cell1.setCellValue("Failure");
					}
				}
				InputStream is = null;
				ByteArrayOutputStream bos = null;
				try {
					bos = new ByteArrayOutputStream();
					System.out.println("Before Workbook Write");
					workbook.write(bos);
					System.out.println("After Workbook Write");
					byte[] barray = bos.toByteArray();
					is = new ByteArrayInputStream(barray);
					String docTitle = doc.get_Name();
					System.out.println("Before Folderset");
					FolderSet folderSet = doc.get_FoldersFiledIn();
					System.out.println("After Folderset");
					Folder folder = null;
					Iterator<Folder> folderSetIterator = folderSet.iterator();
					if (folderSetIterator.hasNext()) {
						folder = folderSetIterator.next();
					}
					String folderPath = folder.get_PathName();
					folderPath += " Response";
					System.out.println("Folder path" + folderPath);
					Folder responseFolder = Factory.Folder.fetchInstance(os, folderPath, null);
					System.out.println("Before Document Save");
					Document updateDoc = updateDocument(os, is, doc, docTitle);
					System.out.println("After Document Save");
					ReferentialContainmentRelationship rc = responseFolder.file(updateDoc, AutoUniqueName.AUTO_UNIQUE,
							docTitle, DefineSecurityParentage.DO_NOT_DEFINE_SECURITY_PARENTAGE);
					rc.save(RefreshMode.REFRESH);

				} catch (Exception e) {
					System.out.println(e);
					e.printStackTrace();
				} finally {
					if (bos != null) {
						bos.close();
					}
					if (is != null) {
						is.close();
					}
					if (stream != null) {
						stream.close();
					}
				}
			}
		} catch (Exception e) {
			System.out.println(e);
			e.printStackTrace();
			throw new RuntimeException(e);
		} finally {
			CaseMgmtContext.set(origCmctx);
		}
	}

	private Document updateDocument(ObjectStore os, InputStream is, Document doc, String docTitle) {
		// TODO Auto-generated method stub
		String docClassName = doc.getClassName() + "Response";
		Document updateDoc = Factory.Document.createInstance(os, docClassName);
		ContentElementList contentList = Factory.ContentElement.createList();
		ContentTransfer contentTransfer = Factory.ContentTransfer.createInstance();
		contentTransfer.setCaptureSource(is);
		contentTransfer.set_RetrievalName(docTitle + ".xlsx");
		contentTransfer.set_ContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
		contentList.add(contentTransfer);

		updateDoc.set_ContentElements(contentList);
		updateDoc.checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MAJOR_VERSION);
		Properties p = updateDoc.getProperties();
		p.putValue("DocumentTitle", docTitle);
		updateDoc.setUpdateSequenceNumber(null);
		updateDoc.save(RefreshMode.REFRESH);
		return updateDoc;
	}

	private static Object getCharValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC:
			return cell.getNumericCellValue();

		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		}
		return null;
	}
}
