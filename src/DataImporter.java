import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.*;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;

/**
 * Imports data from spreadsheets into the FreelancerDB.
 *
 * NOTE: Drop all tables and create them again before running this program.
 */
public class DataImporter {
    // Connection variables
    private static DatabaseConnectionService dbService = new DatabaseConnectionService("golem.csse.rose-hulman.edu","FreelanceDB");
    private static String dbUser = "kangd2";
    private static String dbPass = "csse333";

    // Datapaths to spreadsheets
    private static final String clientsDatapath = "C:\\Users\\kangd2\\Downloads\\Clients.xlsx";
    private static final String freelancersDatapath = "C:\\Users\\kangd2\\Downloads\\Freelancers.xlsx";
    private static final String jobsDatapath = "C:\\Users\\kangd2\\Downloads\\Jobs.xlsx";
    private static final String contactInformationDatapath = "C:\\Users\\kangd2\\Downloads\\Contact_Information.xlsx";
    private static final String industryDatapath = "C:\\Users\\kangd2\\Downloads\\Industry.xlsx";
    private static final String applicationsDatapath = "C:\\Users\\kangd2\\Downloads\\Applications.xlsx";

    public static void main(String[] args) {
        dbService.connect(dbUser, dbPass);

        Map<String, String> dictionary = new LinkedHashMap<>();
        List<Map<String, Object>> parsedData = null;

        // Populate User table
        System.out.println("=============================================");
        System.out.println("Populating User Table");
        System.out.println();
        dictionary.clear();
        dictionary.put("FName", "FirstName");
        dictionary.put("LName", "LastName");
        dictionary.put("Email", "Email");
        dictionary.put("Password Salt", "Salt");
        dictionary.put("Password Hash", "Hash");
        dictionary.put("Bio", "Bio");
        toggleIdentityInsert("User", true);
        parsedData = parseData(clientsDatapath, dictionary);
        for (Map<String, Object> rowData : parsedData) {
            addUser((String) rowData.get("FirstName"), (String) rowData.get("LastName"),
                    (String) rowData.get("Email"), (String) rowData.get("Salt"), (String) rowData.get("Hash"), (String) rowData.get("Bio"));
        }
        parsedData = parseData(freelancersDatapath, dictionary);
        for (Map<String, Object> rowData : parsedData) {
            addUser((String) rowData.get("FirstName"), (String) rowData.get("LastName"),
                    (String) rowData.get("Email"), (String) rowData.get("Salt"), (String) rowData.get("Hash"), (String) rowData.get("Bio"));
        }
        toggleIdentityInsert("User", false);
        System.out.println();

        // Populate Clients table
        System.out.println("=============================================");
        System.out.println("Populating Clients Table");
        System.out.println();
        dictionary.clear();
        dictionary.put("Email", "UserEmail");
        dictionary.put("Cardholder Name", "NameOnCard");
        dictionary.put("Expiration", "CardExpiration");
        dictionary.put("Card Number", "CardNumber");
        dictionary.put("CVV", "CVV");
        toggleIdentityInsert("Clients", true);
        parsedData = parseData(clientsDatapath, dictionary);
        for (Map<String, Object> rowData : parsedData) {
            addClient((String) rowData.get("UserEmail"), (String) rowData.get("NameOnCard"),
                    (String) rowData.get("CardExpiration"), (String) rowData.get("CardNumber"), ((Double) rowData.get("CVV")).intValue());
        }
        toggleIdentityInsert("Clients", false);
        System.out.println();

        // Populate Freelancer table
        System.out.println("=============================================");
        System.out.println("Populating Freelancer Table");
        System.out.println();
        dictionary.clear();
        dictionary.put("Email", "UserEmail");
        dictionary.put("Hourly Rate", "HourlyRate");
        dictionary.put("Skills", "Skill Description");
        dictionary.put("Routing Number", "RoutingNumber");
        dictionary.put("Account Number", "AccountNumber");
        toggleIdentityInsert("Freelancer", true);
        parsedData = parseData(freelancersDatapath, dictionary);
        for (Map<String, Object> rowData : parsedData) {
            addFreelancer((String) rowData.get("UserEmail"), ((Double) rowData.get("HourlyRate")).intValue(),
                    (String) rowData.get("Skill Description"), ((Double) rowData.get("RoutingNumber")).intValue(), ((Double) rowData.get("AccountNumber")).intValue());
        }
        toggleIdentityInsert("Freelancer", false);
        System.out.println();

        // Populate Job table
        System.out.println("=============================================");
        System.out.println("Populating Job Table");
        System.out.println();
        dictionary.clear();
        dictionary.put("Title", "Title");
        dictionary.put("Description", "Description");
        dictionary.put("Budget", "Budget");
        dictionary.put("PostedBy", "PosterMode");
        dictionary.put("AcceptID", "PartnerEmail");
        dictionary.put("AuthorID", "PosterEmail");
        toggleIdentityInsert("Job", true);
        parsedData = parseData(jobsDatapath, dictionary);
        for (Map<String, Object> rowData : parsedData) {
            addJob((String) rowData.get("Title"), (String) rowData.get("Description"), (Double) rowData.get("Budget"),
                    (String) rowData.get("PosterMode"), (String) rowData.get("PartnerEmail"), (String) rowData.get("PosterEmail"));

        }
        toggleIdentityInsert("Job", false);
        System.out.println();

        // Populate ContactInformation table
        System.out.println("=============================================");
        System.out.println("Populating ContactInformation Table");
        System.out.println();
        dictionary.clear();
        dictionary.put("UserID", "UserEmail");
        dictionary.put("URL", "URL");
        dictionary.put("Platform", "Platform");
        toggleIdentityInsert("ContactInformation", true);
        parsedData = parseData(contactInformationDatapath, dictionary);
        for (Map<String, Object> rowData : parsedData) {
            addContactInfo((String) rowData.get("UserEmail"), (String) rowData.get("URL"), (String) rowData.get("Platform"));
        }
        toggleIdentityInsert("ContactInformation", false);
        System.out.println();

        // Populate Industry table
        System.out.println("=============================================");
        System.out.println("Populating Industry Table");
        System.out.println();
        dictionary.clear();
        dictionary.put("IndustryName", "IndustryName");
        toggleIdentityInsert("Industry", true);
        parsedData = parseData(industryDatapath, dictionary);
        for (Map<String, Object> rowData : parsedData) {
            addIndustry((String) rowData.get("IndustryName"));
        }
        toggleIdentityInsert("Industry", false);
        System.out.println();

        // Populate Applies table
        System.out.println("=============================================");
        System.out.println("Populating Applies Table");
        System.out.println();
        dictionary.clear();
        dictionary.put("ApplicantID", "UserEmail");
        dictionary.put("JobID", "JobID");
        toggleIdentityInsert("Applies", true);
        parsedData = parseData(applicationsDatapath, dictionary);
        for (Map<String, Object> rowData : parsedData) {
            addApplication((String) rowData.get("UserEmail"), ((Double) rowData.get("JobID")).intValue());
        }
        toggleIdentityInsert("Applies", false);
        System.out.println();

        // Populate UserInvolvedIn table
        System.out.println("=============================================");
        System.out.println("Populating UserInvolvedIn Table");
        System.out.println();
        dictionary.clear();
        dictionary.put("Industry", "IndustryName");
        dictionary.put("Email", "UserEmail");
        toggleIdentityInsert("UserInvolvedIn", true);
        parsedData = parseData(clientsDatapath, dictionary);
        for (Map<String, Object> rowData : parsedData) {
            addUserIndustry((String) rowData.get("IndustryName"), (String) rowData.get("UserEmail"));
        }
        parsedData = parseData(freelancersDatapath, dictionary);
        for (Map<String, Object> rowData : parsedData) {
            addUserIndustry((String) rowData.get("IndustryName"), (String) rowData.get("UserEmail"));
        }
        toggleIdentityInsert("UserInvolvedIn", false);
        System.out.println();

        // Populate JobInvolvedIn table
        System.out.println("=============================================");
        System.out.println("Populating JobInvolvedIn Table");
        System.out.println();
        dictionary.clear();
        dictionary.put("Industry", "IndustryName");
        dictionary.put("JobID", "JobID");
        toggleIdentityInsert("JobInvolvedIn", true);
        parsedData = parseData(jobsDatapath, dictionary);
        for (Map<String, Object> rowData : parsedData) {
            addJobIndustry((String) rowData.get("IndustryName"), ((Double) rowData.get("JobID")).intValue());
        }
        toggleIdentityInsert("JobInvolvedIn", false);
        System.out.println();

        dbService.closeConnection();
    }

    public static boolean toggleIdentityInsert(String table, boolean identityInsert) {
        try (PreparedStatement statement = dbService.getConnection().prepareStatement("SET IDENTITY_INSERT dbo.? ?;")) {
            statement.setString(1, table);
            if (identityInsert) {
                statement.setString(2, "ON");
            } else {
                statement.setString(2, "OFF");
            }
            statement.executeQuery();
        }
        catch (SQLException e) {
            return false;
        }

        return true;
    }

    public static boolean addApplication(String email, int job) {
        try (CallableStatement proc = dbService.getConnection().prepareCall("{ ? = call dbo.AddApplication(?, ?) }")) {
            proc.registerOutParameter(1, Types.INTEGER);
            proc.setInt(2, job);
            proc.setString(3, email);
            proc.execute();
            int returnValue = proc.getInt(1);
            if (returnValue != 0) {
                return false;
            }
        }
        catch (SQLException e) {
            System.err.println(e);
            return false;
        }

        return true;
    }

    public static boolean addJobIndustry(String industry, int job) {
        try (CallableStatement proc = dbService.getConnection().prepareCall("{ ? = call dbo.AddIndustrytoJob(?, ?) }")) {
            proc.registerOutParameter(1, Types.INTEGER);
            proc.setString(2, industry);
            proc.setInt(3, job);
            proc.execute();
            int returnValue = proc.getInt(1);
            if (returnValue != 0) {
                return false;
            }
        }
        catch (SQLException e) {
            System.err.println(e);
            return false;
        }

        return true;
    }

    public static boolean addUserIndustry(String industry, String email) {
        try (CallableStatement proc = dbService.getConnection().prepareCall("{ ? = call dbo.AddIndustrytoUser(?, ?) }")) {
            proc.registerOutParameter(1, Types.INTEGER);
            proc.setString(2, email);
            proc.setString(3, industry);
            proc.execute();
            int returnValue = proc.getInt(1);
            if (returnValue != 0) {
                return false;
            }
        }
        catch (SQLException e) {
            System.err.println(e);
            return false;
        }

        return true;
    }

    public static boolean addIndustry(String name) {
        try (CallableStatement proc = dbService.getConnection().prepareCall("{ ? = call dbo.AddIndustry(?) }")) {
            proc.registerOutParameter(1, Types.INTEGER);
            proc.setString(2, name);
            proc.execute();
            int returnValue = proc.getInt(1);
            if (returnValue != 0) {
                return false;
            }
        }
        catch (SQLException e) {
            System.err.println(e);
            return false;
        }

        return true;
    }

    public static boolean addContactInfo(String email, String url, String platform) {
        try (CallableStatement proc = dbService.getConnection().prepareCall("{ ? = call dbo.AddContactInformation(?, ?, ?) }")) {
            proc.registerOutParameter(1, Types.INTEGER);
            proc.setString(2, email);
            proc.setString(3, url);
            proc.setString(4, platform);
            proc.execute();
            int returnValue = proc.getInt(1);
            if (returnValue != 0) {
                return false;
            }
        }
        catch (SQLException e) {
            System.err.println(e);
            return false;
        }

        return true;
    }

    public static boolean addFreelancer(String email, double rate, String skills, int routing, int account) {
        try (CallableStatement proc = dbService.getConnection().prepareCall("{ ? = call dbo.AddFreelancer(?, ?, ?, ?, ?) }")) {
            proc.registerOutParameter(1, Types.INTEGER);
            proc.setString(2, email);
            proc.setDouble(3, rate);
            proc.setString(4, skills);
            proc.setInt(5, routing);
            proc.setInt(6, account);
            proc.execute();
            int returnValue = proc.getInt(1);
            if (returnValue != 0) {
                return false;
            }
        }
        catch (SQLException e) {
            System.err.println(e);
            return false;
        }

        return true;
    }

    public static boolean addClient(String email, String nameOnCard, String expiration, String cardNumber, int cvv) {
        try (CallableStatement proc = dbService.getConnection().prepareCall("{ ? = call dbo.AddClient(?, ?, ?, ?, ?) }")) {
            DateFormat df = new SimpleDateFormat("MM-dd-yyyy");
            java.sql.Date expirationDate = new java.sql.Date(df.parse(expiration).getTime());

            proc.registerOutParameter(1, Types.INTEGER);
            proc.setString(2, email);
            proc.setString(3, nameOnCard);
            proc.setDate(4, expirationDate);
            proc.setString(5, cardNumber);
            proc.setInt(6, cvv);
            proc.execute();
            int returnValue = proc.getInt(1);
            if (returnValue != 0) {
                return false;
            }
        }
        catch (SQLException e) {
            System.err.println(e);
            return false;
        } catch (ParseException e) {
            System.err.println(e);
            return false;
        }

        return true;
    }

    public static boolean addJob(String title, String desc, double budget, String mode, String partnerEmail, String posterEmail) {
        try (CallableStatement proc = dbService.getConnection().prepareCall("{ ? = call dbo.AddJob(?, ?, ?, ?, ?, ?, ?) }")) {
            proc.registerOutParameter(1, Types.INTEGER);
            proc.setString(2, title);
            proc.setString(3, desc);
            proc.setDouble(4, budget);
            proc.setString(5, mode);
            if (partnerEmail == null) {
                proc.setNull(6, Types.NVARCHAR);
            } else {
                proc.setString(6, partnerEmail);
            }
            proc.setString(7, posterEmail);
            proc.registerOutParameter(8, Types.INTEGER);
            proc.execute();
            int returnValue = proc.getInt(1);
            if (returnValue != 0) {
                return false;
            }
        }
        catch (SQLException e) {
            System.err.println(e);
            return false;
        }

        return true;
    }

    public static boolean addUser(String firstName, String lastName, String email, String salt, String hash, String bio) {
        try (CallableStatement proc = dbService.getConnection().prepareCall("{ ? = call dbo.AddUser(?, ?, ?, ?, ?, ?) }")) {
            proc.registerOutParameter(1, Types.INTEGER);
            proc.setString(2, firstName);
            proc.setString(3, lastName);
            proc.setString(4, email);
            proc.setString(5, salt);
            proc.setString(6, hash);
            proc.setString(7, bio);
            proc.execute();
            int returnValue = proc.getInt(1);
            if (returnValue != 0) {
                return false;
            }
        }
        catch (SQLException e) {
            System.err.println(e);
            return false;
        }

        return true;
    }


    /**
     * Parses an Excel file row by row
     *
     * @param path: path to file
     * @param dictionary: defines which columns are needed and what they are named in the SQL server
     * @return a list of HashMaps that pair each value to their respective correct header
     */
    // Source: https://www.javatpoint.com/how-to-read-excel-file-in-java
    public static List<Map<String, Object>> parseData(String path, Map<String, String> dictionary) {
        List<Map<String, Object>> parsedData = new ArrayList<>();

        try {
            File file = new File(path);   //creating a new file instance
            FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file

            //creating Workbook instance that refers to .xlsx file
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object
            Iterator<Row> itr = sheet.iterator();    //iterating over excel file

            FormulaEvaluator formulaEvaluator=wb.getCreationHelper().createFormulaEvaluator();

            Row row = itr.next();
            Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column

            // Data structures used to parse table data
            List<String> header = new ArrayList<>();

            // Parse header
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                header.add(cell.getStringCellValue());
            }

            // Parse each row
            while (itr.hasNext()) {
                row = itr.next();
                cellIterator = row.cellIterator();   //iterating over each column

                Map<String, Object> rowData = new HashMap<>();

                // Parse each cell
                for (int i = 0; i < header.size(); i++) {
                    Cell cell = cellIterator.next();

                    String curHeader = header.get(i);
                    if (!dictionary.containsKey(curHeader)) {
                        // If not needed, skip this cell
                        continue;
                    }

                    String newHeader = dictionary.get(curHeader);
                    Object data;

                    switch(formulaEvaluator.evaluateInCell(cell).getCellType())
                    {
                        case Cell.CELL_TYPE_BOOLEAN:
                            data = cell.getBooleanCellValue();
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            data = cell.getNumericCellValue();
                            break;
                        case Cell.CELL_TYPE_STRING:
                            data = cell.getStringCellValue();
                            if (!data.equals("null")) {
                                break;
                            }
                        default:
                            data = null;
                    }
                    rowData.put(newHeader, data);
                }
                parsedData.add(rowData);
            }
        } catch(Exception e) {
            e.printStackTrace();
        }

        return parsedData;
    }
}
