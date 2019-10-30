package com.generic.util;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import com.generic.page.LoadTestResult;
import com.generic.setup.EnvironmentFiles;
import com.generic.setup.ReleaseDetails;

public class SqLiteUtils {

	public String[] PageName = new String[100];
	public String[] PageHits = new String[100];
	public String[] Onload = new String[100];

	public static void closeConnection(Connection conn) {
		//// getCurrentFunctionName(true);
		if (conn != null) {
			try {
				conn.close();
			} catch (SQLException e) {
				System.out.println("connection closed");
			}
		}
		//// getCurrentFunctionName(false);
	}// close connection

	public static void insertCBIRelease(String TableName, String DatabaseName) {
		String url = "jdbc:sqlite:" + EnvironmentFiles.getDatabasePath() + "\\" + DatabaseName;

		try (Connection conn = DriverManager.getConnection(url)) {

			String isRowExists = "SELECT count(*) FROM cbi_releases where release_name = " + "'"
					+ ReleaseDetails.getReleaseName() + "' and release_run_num = '" + ReleaseDetails.getRunNum()
					+ "' ;";

			PreparedStatement isRowExistsStatement = conn.prepareStatement(isRowExists);
			ResultSet rs1 = isRowExistsStatement.executeQuery();

			System.out.println("~~~~IsRowExists:" + rs1.getInt("count(*)"));

			if (rs1.getInt("count(*)") == 0) {
				String Statement = "INSERT INTO " + TableName + " (release_name, release_run_num) " + " VALUES(" + "'"
						+ ReleaseDetails.getReleaseName() + "'," + "'" + ReleaseDetails.getRunNum() + "'" + ")";

				System.out.println(Statement);

				PreparedStatement ps = conn.prepareStatement(Statement);

				// if it returns less than 0, no rows were inserted
				if (ps.executeUpdate() > 0)
					System.out.println("Inserted sucessfuly");
				else
					System.out.println("Failed to insert");
				closeConnection(conn);
			}

		} catch (SQLException e) {
			System.out.println(e.getMessage());
		}
	}// insertCBIRelease

	public static void insertCBIBrandsPages(LoadTestResult details, String TableName, String DatabaseName) {
		String url = "jdbc:sqlite:" + EnvironmentFiles.getDatabasePath() + "\\" + DatabaseName;

		try (Connection conn = DriverManager.getConnection(url)) {

			String isRowExists = "SELECT count(*) FROM cbi_brands_pages where page_name = " + "'"
					+ details.getLoadTestDetails().pageName + "' and brand = '" + details.getLoadTestDetails().Brand
					+ "' ;";

			PreparedStatement isRowExistsStatement = conn.prepareStatement(isRowExists);
			ResultSet rs1 = isRowExistsStatement.executeQuery();

			System.out.println("~~~~IsRowExists:" + rs1.getInt("count(*)"));

			if (rs1.getInt("count(*)") == 0) {
				String Statement = "INSERT INTO " + TableName + " (PAGE_NAME, BRAND) " + " VALUES(" + "'"
						+ details.getLoadTestDetails().pageName + "'," + "'" + details.getLoadTestDetails().Brand + "'"
						+ ")";

				System.out.println(Statement);

				PreparedStatement ps = conn.prepareStatement(Statement);

				// if it returns less than 0, no rows were inserted
				if (ps.executeUpdate() > 0)
					System.out.println("Inserted sucessfuly");
				else
					System.out.println("Failed to insert");
				closeConnection(conn);
			}

		} catch (SQLException e) {
			System.out.println(e.getMessage());
		}
	}// insertCBIBrandsPages

	public static void insertLoadTestResult(LoadTestResult details, String TableName, String DatabaseName) {
		String url = "jdbc:sqlite:" + EnvironmentFiles.getDatabasePath() + "\\" + DatabaseName;

		try (Connection conn = DriverManager.getConnection(url)) {

			String getPageId = "SELECT page_id FROM cbi_brands_pages where page_name = " + "'"
					+ details.getLoadTestDetails().pageName + "' and brand = '" + details.getLoadTestDetails().Brand
					+ "' ;";

			String getReleaseId = "SELECT release_id FROM cbi_releases where release_name = " + "'"
					+ ReleaseDetails.getReleaseName() + "' ;";

			System.out.println(getPageId);
			System.out.println(getReleaseId);

			PreparedStatement getPageIdStatement = conn.prepareStatement(getPageId);
			ResultSet rs1 = getPageIdStatement.executeQuery();

			PreparedStatement getReleaseIdStatement = conn.prepareStatement(getReleaseId);
			ResultSet rs2 = getReleaseIdStatement.executeQuery();

			String isRowExists = "SELECT count(*) FROM " + TableName + " where page_id = " + rs1.getInt("page_id")
					+ " and release_id = " + rs2.getInt("release_id") + " ;";

			System.out.println(isRowExists);

			PreparedStatement isRowExistsStatement = conn.prepareStatement(isRowExists);
			ResultSet exists = isRowExistsStatement.executeQuery();

			System.out.println("~~~~IsRowExists:" + exists.getInt("count(*)"));

			if (exists.getInt("count(*)") == 0) {

				String Statement = "INSERT INTO " + TableName + " (page_id, release_id, brand, page_hits, onload) "
						+ " VALUES(" + rs1.getInt("page_id") + ',' + rs2.getInt("release_id") + ", '"
						+ details.getLoadTestDetails().Brand + "' ,'" + details.getLoadTestDetails().pageHits + "' ,'"
						+ details.getLoadTestDetails().onload + "' );";

				System.out.println(Statement);

				PreparedStatement ps = conn.prepareStatement(Statement);
				// if it returns less than 0, no rows were inserted
				if (ps.executeUpdate() > 0)
					System.out.println("Inserted sucessfuly");
				else
					System.out.println("Failed to insert");
				closeConnection(conn);
			}
		} catch (SQLException e) {
			System.out.println(e.getMessage());
		}
	}// insertBrandsLoadTestResult

	public void getReleaseResults(String release_name, String TableName, String DatabaseName) {
		String url = "jdbc:sqlite:" + EnvironmentFiles.getDatabasePath() + "\\" + DatabaseName;
		ResultSet exists = null;

		try (Connection conn = DriverManager.getConnection(url)) {

			String statement = "SELECT p.page_name, d.page_hits, d.onload FROM " + TableName
					+ " d, cbi_brands_pages p, cbi_releases r where d.page_id = p.page_id "
					+ "AND d.release_id = r.release_id AND r.release_name = '" + release_name
					+ "' order by p.page_name;";

			System.out.println(statement);

			PreparedStatement rs = conn.prepareStatement(statement);
			exists = rs.executeQuery();

			int i = 0;
			while (exists.next()) {
				PageName[i] = exists.getString("page_name");
				PageHits[i] = exists.getString("page_hits");
				Onload[i] = exists.getString("onload");
				i++;
			}

		} catch (SQLException e) {
			System.out.println(e.getMessage());
		}
	}// getReleaseResults
}
