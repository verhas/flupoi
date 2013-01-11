package javax0.flupoi;

import java.io.IOException;
import java.io.InputStream;
import java.util.Collection;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public interface Workbook {

	/**
	 * Read the content of the workbook from the specified stream.
	 * 
	 * @param is
	 *            the input stream where the XLS or XLSX file is.
	 * 
	 */
	Workbook stream(InputStream is);

	/**
	 * Select the sheet from the workbook for the successive operations.
	 * 
	 * @param name
	 *            the name of the sheet
	 */
	Workbook sheet(String name);

	/**
	 * Specify the range for the successive operations.
	 * 
	 * @param range
	 *            specifies the range to work. The format of the range can be:
	 *            <ul>
	 *            <li><tt>Nx</tt>, e.g.: <tt>A5</tt> specifies a single cell
	 *            range.
	 *            <li><tt>Nx:My</tt>, e.g.: <tt>A5:C7</tt> specifies a range
	 *            exactly. Such a range specification should be followed by
	 *            {@link #up()}, {@link #down()}, {@link #left()} or
	 *            {@link #right()} method call.
	 *            <li><tt>N*:My</tt>, e.g.: <tt>A*:C7</tt> specifies a range
	 *            that starts with the row <tt>A7
	 *            - C7</tt> and goes down to <tt>A0 - C0</tt> or until range
	 *            condition stops the range processing
	 *            <li><tt>Ny:M*</tt>, e.g.: <tt>A5:C*</tt> specifies a range
	 *            that starts with the row <tt>A5
	 *            - C5</tt> and goes up to the end of the table or until range
	 *            condition stops the range processing
	 *            <li><tt>*x:My</tt>, e.g.: <tt>*5:C7</tt> specifies a range
	 *            that starts with the column <tt>C5:C7</tt> and goes left to
	 *            A5:A7</tt> or until range condition stops the range processing
	 *            <li><tt>Nx,*y</tt>, e.g.: <tt>A5:*7</tt> specifies a range
	 *            that starts with the column <tt>A5:A7</tt> and goes to the
	 *            right until the end of the table or until the condition stops
	 *            the range processing
	 *            </ul>
	 *            Note that any range end point can be write in the form
	 *            <tt>(n,m)</tt>, for example you can write <tt>(0,0)</tt>
	 *            instead of <tt>A0</tt>. The range <tt>A5:C7</tt> can be
	 *            written as <tt>(0,5):(2,7)</tt>.
	 *            <p>
	 *            You should NOT use <tt>+</tt> or </tt>-</tt> in front of the
	 *            numbers in this case, because that will mean relative cell
	 *            coordinate based on the last range or position.
	 * 
	 */
	Workbook range(String range);

	/**
	 * Specifies that the last specified range is to be processed from the
	 * lowest row upward towards the smaller numeric indexes.
	 */
	Workbook up();

	/**
	 * Specifies that the last specified range is to be processed from the
	 * highest (smallest index) row downwards.
	 */
	Workbook down();

	/**
	 * Specifies that the range is to be processed leftward.
	 */
	Workbook left();

	/**
	 * Specifies that the range is to be processed from left to right.
	 * 
	 * @return
	 */
	Workbook right();

	/**
	 * Specifies the condition that stops the processing of the range.
	 * 
	 * @param condition
	 */
	Workbook until(Condition condition);

	/**
	 * Specifies that the range is stopped by the first line that matches the
	 * string.
	 * 
	 * @param match
	 */
	Workbook until(String match);

	/**
	 * Specifies that the range is stopped by the first empty line
	 */
	Workbook untilEmptyLine();

	/**
	 * Specifies that the range is stopped by the first empty cell that is on
	 * the first column of the range or on the first row of the range.
	 */
	Workbook untilEmptyCell();

	/**
	 * Specifies the column or raw names.
	 * 
	 * @param names
	 */
	Workbook names(String[] names);

	/**
	 * Same as {@link #names(String[])}, except the names are given in a single
	 * string
	 * 
	 * @param names
	 *            the comma separated names
	 */
	Workbook names(String names);

	/**
	 * Specify the target for the range.
	 * 
	 * @param key
	 *            key, a String or some presumably immutable object appropriate
	 *            to be used as a key in a map. This key should be used later as
	 *            parameter in the method {@link #fetch(Object)}.
	 * @param type
	 *            the type of the class to convert each row/column of the range.
	 */
	Workbook target(Object key, Class<?> type);

	/**
	 * Specify to move the last processed point to which a range can be
	 * specified relative. Note that both movements, horizontal and vertical can
	 * be positive and negative.
	 * <p>
	 * The final position will never be negative.
	 * 
	 * @param x
	 *            the number of cells to move horizontal
	 * @param y
	 *            the number of cells to move vertical
	 */
	Workbook skip(int x, int y);

	/**
	 * Defines a single cell range with a name.
	 * 
	 * @param name
	 * @param range
	 */
	Workbook cell(String name, String range);

	/**
	 * Execute the conversion.
	 */
	Workbook execute() throws InvalidFormatException, IOException,
			InstantiationException, IllegalAccessException,
			NoSuchFieldException, SecurityException;

	/**
	 * Fetch the object array that was stored in the target identified by the
	 * key.
	 * 
	 * @param key
	 * @return the collection that was gathered for the specific key
	 */
	Collection<?> fetch(Object key);
}
