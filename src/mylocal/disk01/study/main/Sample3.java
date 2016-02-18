package mylocal.disk01.study.main;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class Sample3 {

	private static final String DRIVER_CLASS = "org.mariadb.jdbc.Driver";
	private static final String DB_URL = "jdbc:mariadb://localhost:3306/mydb";
	private static final String DB_USER = "root";
	private static final String DB_PASS = "maria";

	private static final String SAMPLE_SQL1 = "select id, name from os_mst;";

	private final Logger logger = LoggerFactory.getLogger(this.getClass());

	private Connection getConnection() {

		try {
			Class.forName(DRIVER_CLASS);

			Connection conn = DriverManager.getConnection(DB_URL, DB_USER, DB_PASS);

			try {
				conn.setAutoCommit(false);
				return conn;
			} catch (SQLException e) {
				// TODO: handle exception
				close(conn);
				logger.error("That Error occurred when connecting to the DB. ",e);
				throw new RuntimeException("That Error occurred when connecting to the DB.", e);
			}
		} catch (ClassNotFoundException e) {
			// TODO: handle exception
			logger.error("Cloud not find the JDBC Driver." + DRIVER_CLASS, e);
			throw new RuntimeException("Cloud not find the JDBC Driver." + DRIVER_CLASS, e);
		} catch (SQLException e) {
			// TODO: handle exception
			logger.error("That Error occurred when connecting to the DB. ",e);
			throw new RuntimeException("That Error occurred when connecting to the DB.", e);
		}
	}

	private void close(Connection conn) {
		if (conn == null) {
			return;
		}
		try {
			if (!conn.isClosed()) {
				conn.close();
			}
		} catch (SQLException e) {
			// TODO: handle exception
			logger.warn(e.getMessage(), e);
		}

	}

	private void close(Statement statement) {
		if (statement == null) {
			return;
		}
		try {
			statement.close();
		} catch (SQLException e) {
			// TODO: handle exception
			logger.warn(e.getMessage(), e);
		}
	}


	private void close(ResultSet rs) {
		if (rs == null) {
			return;
		}
		try {
			rs.close();
		} catch (SQLException e) {
			// TODO: handle exception
			logger.warn(e.getMessage(), e);
		}
	}

	public static void main(String[] args) {
		// TODO 自動生成されたメソッド・スタブ
		new Sample3().execute();
	}

	private void execute() {

		Connection conn = null;
		Statement  st = null;
		ResultSet  rs = null;

		try {
			conn = getConnection();
			st = conn.createStatement();
			rs = st.executeQuery(SAMPLE_SQL1);

			boolean hashNext = rs.next();
			if (!hashNext) {
				logger.warn("no record!");
				return;
			}

			System.out.println(
					rs.getInt(1)
					+ "," + rs.getString(2)
			);

			while (rs.next()) {
				System.out.println(
						rs.getInt(1)
						+ "," + rs.getString(2)
				);
			}
		} catch (RuntimeException e) {
			// TODO: handle exception
		} catch (SQLException e) {
			// TODO: handle exception
			logger.error("That Error occurred when access into the DB. ",e);
		} finally {
			// TODO: handle finally clause
			close(rs);
			close(st);
			close(conn);
		}

	}

}
