package routines.system;

import java.sql.SQLException;

public class TalendDataSource {

	private final javax.sql.DataSource ds;
	private java.sql.Connection conn;

	public TalendDataSource(javax.sql.DataSource ds) {
		this.ds = ds;
	}

	public java.sql.Connection getConnection() throws SQLException {
		if (null == conn) {
			conn = ds.getConnection();
		}
		return conn;
	}

	public javax.sql.DataSource getRawDataSource() {
		return ds;
	}

	public void close() throws SQLException {
		if (null != conn) {
			conn.close();
			conn = null;
		}
	}
}
