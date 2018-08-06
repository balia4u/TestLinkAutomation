package com.bigBasket.uploadTestCases;

import org.w3c.dom.CDATASection;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.internet.MimeBodyPart;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

public class ModifyXMLFile {

	public String[][] getTestCases() {
		String TestCasefilepath = System.getProperty("user.dir") + "//TestData//" + System.getProperty("fileName");
		ExcelLibrary exlb = new ExcelLibrary(TestCasefilepath);
		int totalRow = 0, toatlColumn = 0, i = 0, j = 0;
		String[][] data = null;
		try {
			totalRow = exlb.getRowCount();
			toatlColumn = exlb.getColumnCount();
			data = new String[totalRow + 1][toatlColumn];
			for (i = 0; i <= totalRow; i++) {
				for (j = 0; j < toatlColumn; j++) {
					data[i][j] = exlb.getExcelData(i, j);
				}
			}
		} catch (Exception e) {
			System.out.println("Error while getting row " + i + " and column " + j + " data");
			e.printStackTrace();
		}
		return data;
	}

	public static void main(String args[]) {
		ModifyXMLFile mfx3 = new ModifyXMLFile();
		String[][] data = mfx3.getTestCases();
		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
		DocumentBuilder dBuilder;
		try {
			dBuilder = dbFactory.newDocumentBuilder();
			Document doc = dBuilder.newDocument();
			Element rootElementTestCases = doc.createElement("testcases");
			doc.appendChild(rootElementTestCases);
			int totalTestCase = 0;
			for (int i = 0; i < data.length; i++) {
				if ((i == 0) || (!(data[i][0] != null)) || (data[i][0].isEmpty()))
					continue;
				Node importedNodeTestCase = mfx3.createTestCaseNode(data[i][0], data[i][1], data[i][2], data[i][3],
						data[i][4], data[i][5], data[i][6], data[i][7], data[i][8], data[i][9], i);
				rootElementTestCases.appendChild(doc.importNode(importedNodeTestCase, true));
				totalTestCase++;
			}
			System.out.println("Total test cases " + totalTestCase);
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			DOMSource source = new DOMSource(doc);
			StreamResult file = new StreamResult(
					new File(System.getProperty("user.dir") + "//TestData//testcase1.xml"));
			transformer.transform(source, file);
			Thread.sleep(1000);
			mfx3.mailService();
			System.out.println("DONE");

		} catch (ParserConfigurationException pce) {
			pce.printStackTrace();
		} catch (TransformerException tfe) {
			tfe.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	@SuppressWarnings("null")
	public Node createTestCaseNode(String name, String summary, String preconditions, String steps,
			String expectedresults, String JiraTicket, String Automation_Complexity, String Is_Automated,
			String Manual_Execution_Complexity, String Search_Text, int row) {
		String filepath;
		DocumentBuilderFactory docFactory;
		DocumentBuilder docBuilder;
		Document doc = null;
		Node testcase = null;
		try {
			filepath = System.getProperty("user.dir") + "//TestData//testcase.xml";
			docFactory = DocumentBuilderFactory.newInstance();
			docBuilder = docFactory.newDocumentBuilder();
			doc = docBuilder.parse(filepath);
			testcase = doc.getElementsByTagName("testcase").item(0);
			NamedNodeMap attr = testcase.getAttributes();
			Node nodeAttr = attr.getNamedItem("name");
			name=name.substring(0, Math.min(name.length(), 100));
			nodeAttr.setTextContent(name);
			NodeList list = testcase.getChildNodes();

			for (int i = 0; i < list.getLength(); i++) {

				Node node = list.item(i);
				if ("node_order".equals(node.getNodeName())) {
					checkLinesandAddPTag("", node, doc);
				} else if ("externalid".equals(node.getNodeName())) {
					checkLinesandAddPTag("", node, doc);
				} else if ("version".equals(node.getNodeName())) {
					checkLinesandAddPTag("", node, doc);
				} else if ("summary".equals(node.getNodeName())) {
					checkLinesandAddPTag(summary, node, doc);
				} else if ("preconditions".equals(node.getNodeName())) {
					checkLinesandAddPTag(preconditions, node, doc);
				} else if ("execution_type".equals(node.getNodeName())) {
					checkLinesandAddPTag("1", node, doc);
				} else if ("importance".equals(node.getNodeName())) {
					checkLinesandAddPTag("2", node, doc);
				} else if ("steps".equals(node.getNodeName())) {
					NodeList list2 = node.getChildNodes();
					for (int j = 0; j < list2.getLength(); j++) {
						Node stepsChilNodes = list2.item(j);

						if (stepsChilNodes.getNodeName().equalsIgnoreCase("step")) {

							NodeList list3 = stepsChilNodes.getChildNodes();
							for (int k = 0; k < list3.getLength(); k++) {
								Node stepsChilNodes2 = list3.item(k);
								if ("step_number".equals(stepsChilNodes2.getNodeName())) {

									checkLinesandAddPTag("1", stepsChilNodes2, doc);
								} else if ("actions".equals(stepsChilNodes2.getNodeName())) {

									checkLinesandAddPTag(steps, stepsChilNodes2, doc);
								} else if ("expectedresults".equals(stepsChilNodes2.getNodeName())) {
									checkLinesandAddPTag(expectedresults, stepsChilNodes2, doc);
								} else if ("execution_type".equals(stepsChilNodes2.getNodeName())) {
									checkLinesandAddPTag("1", stepsChilNodes2, doc);
								}
							}
						}
					}

				} else if ("custom_fields".equals(node.getNodeName())) {
					NodeList custom_fields = node.getChildNodes();
					for (int j = 0; j < custom_fields.getLength(); j++) {
						Node stepsChilNodes = custom_fields.item(j);
						if (stepsChilNodes.getNodeName().equalsIgnoreCase("custom_field")) {

							NodeList custom_fields2 = stepsChilNodes.getChildNodes();
							for (int k = 0; k < custom_fields2.getLength(); k++) {
								Node stepsChilNodes2 = custom_fields2.item(k);
								if ("value".equals(stepsChilNodes2.getNodeName())) {
									if ((Automation_Complexity != null)
											&& stepsChilNodes2.getTextContent()
													.equalsIgnoreCase("  Not Automatable ")) {
										if (Automation_Complexity.trim().equalsIgnoreCase("Not Automatable")) {
											checkLinesandAddPTag("  Not Automatable ", stepsChilNodes2, doc);
										} else if (Automation_Complexity.trim().equalsIgnoreCase("Easy")) {
											checkLinesandAddPTag("Easy", stepsChilNodes2, doc);
										} else if (Automation_Complexity.trim().equalsIgnoreCase("Hard")) {
											checkLinesandAddPTag("Hard", stepsChilNodes2, doc);
										}
									} else if ((Is_Automated != null)
											&& stepsChilNodes2.getTextContent().equalsIgnoreCase(" no")) {
										if (Is_Automated.trim().equalsIgnoreCase(" no")) {
											checkLinesandAddPTag("no", stepsChilNodes2, doc);
										} else if (Is_Automated.trim().equalsIgnoreCase("yes")) {
											checkLinesandAddPTag(" yes", stepsChilNodes2, doc);
										}
									} else if ((Manual_Execution_Complexity != null)
											&& stepsChilNodes2.getTextContent().equalsIgnoreCase("Easy ")) {
										if (Manual_Execution_Complexity.trim().equalsIgnoreCase("Easy")) {
											checkLinesandAddPTag("Easy ", stepsChilNodes2, doc);
										} else if (Manual_Execution_Complexity.trim().equalsIgnoreCase("Complex")) {
											checkLinesandAddPTag("Complex ", stepsChilNodes2, doc);
										} 
									} else if ((Search_Text != null)
											&& stepsChilNodes2.getTextContent().equalsIgnoreCase("Search")) {
										checkLinesandAddPTag(Search_Text, stepsChilNodes2, doc);
									} else if ((Search_Text == null)
											&& stepsChilNodes2.getTextContent().equalsIgnoreCase("Search")) {
										checkLinesandAddPTag("", stepsChilNodes2, doc);
									} else if ((JiraTicket != null)
											&& stepsChilNodes2.getTextContent().equalsIgnoreCase("BB")) {
										checkLinesandAddPTag(JiraTicket, stepsChilNodes2, doc);
									} else if ((JiraTicket == null)
											&& stepsChilNodes2.getTextContent().equalsIgnoreCase("BB")) {
										checkLinesandAddPTag("", stepsChilNodes2, doc);
									}

								}

							}
						}

					}
				}
			}
		} catch (ParserConfigurationException pce) {
			pce.printStackTrace();
		} catch (IOException ioe) {
			ioe.printStackTrace();
		} catch (SAXException sae) {
			sae.printStackTrace();
		}
		return doc.getElementsByTagName("testcase").item(0);

	}

	public String checkLinesandAddPTag(String data, Node node, Document doc) {
		String newData;
		if (!(data != null)) {
			newData = "";
		} else
			newData = data;

		Element eElement = null;
		CDATASection dirdata = null;

		node.setTextContent("");
		eElement = (Element) node;
		if (data != null && data.contains("\n")) {
			newData = "";
			for (String s : data.split("[\\n\\r]+")) {
				newData += "<p>\n" + s + "</p>\n";
			}
		}
		dirdata = doc.createCDATASection(newData);
		eElement.appendChild(dirdata);

		return newData;
	}

	public void mailService() {
		try {
			String subject = "Attachment Testlink XML file", consolidatedBody = "PFA";
			String[] emailTo = System.getProperty("emailTo").split(",");

			ArrayList<MimeBodyPart> messageBodyList = new ArrayList<MimeBodyPart>();
			MimeBodyPart messageBodyPart1 = new MimeBodyPart();
			messageBodyPart1.setContent(consolidatedBody, "text/html");
			messageBodyList.add(messageBodyPart1);
			MimeBodyPart messageBodyAttachment = new MimeBodyPart();
			DataSource source = new FileDataSource(System.getProperty("user.dir") + "//TestData//testcase1.xml");
			messageBodyAttachment.setDataHandler(new DataHandler(source));
			messageBodyAttachment.setFileName("testcase1.xml");
			messageBodyList.add(messageBodyAttachment);
			new Mailer().sendMail(emailTo, subject, messageBodyList);
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}