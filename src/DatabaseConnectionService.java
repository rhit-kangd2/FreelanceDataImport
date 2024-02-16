import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class DatabaseConnectionService {
    private final String SampleURL = "jdbc:sqlserver://${dbServer};databaseName=${dbName};user=${user};password={${pass}}";

    private Connection connection = null;

    private String databaseName;
    private String serverName;

    public DatabaseConnectionService(String serverName, String databaseName) {
        //DO NOT CHANGE THIS METHOD
        this.serverName = serverName;
        this.databaseName = databaseName;
    }

    public boolean connect(String user, String pass) {
        //BUILD YOUR CONNECTION STRING HERE USING THE SAMPLE URL ABOVE
        String fullUrl = SampleURL
                .replace("${dbServer}", serverName)
                .replace("${dbName}", databaseName)
                .replace("${user}", user)
                .replace("${pass}", pass);
        fullUrl += ";encrypt=true;trustServerCertificate=true";

        try {
            connection = DriverManager.getConnection(fullUrl);
            return true;
        } catch (SQLException e) {
            System.err.println(e);
        }

        return false;
    }


    public Connection getConnection() {
        return this.connection;
    }

    public void closeConnection() {
        try {
            if (connection != null && !connection.isClosed()) {
                connection.close();
            }
        } catch (SQLException e) {
            System.err.println(e);
        }
    }

}
