package oracle;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;
import org.springframework.jdbc.datasource.DriverManagerDataSource;
import javax.sql.DataSource;

@Configuration
@PropertySource("classpath:application.properties")
public class OracleDBConfig {

    @Value("${oracle.datasource.url}")//sostituire url
    private String url;

    @Value("${oracle.datasource.username}") //sostituire username
    private String username;

    @Value("${oracle.datasource.password}") //sostituire password
    private String password;

    @Value("${oracle.jdbc.OracleDriver}")
    private String driverClassName;


    // Bean per la creazione del DataSource
    public DataSource dataSource() {
        DriverManagerDataSource dataSource = new DriverManagerDataSource();
        dataSource.setUrl(url);
        dataSource.setUsername(username);
        dataSource.setPassword(password);
        dataSource.setDriverClassName(driverClassName);
        return dataSource;
    }
}
