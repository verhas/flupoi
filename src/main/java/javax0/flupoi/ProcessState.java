package javax0.flupoi;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

class ProcessState {
	private InputStream is;

	protected void setIs(InputStream is) {
		this.is = is;
	}

	private Workbook workbook;

	private Workbook getWorkbook() throws InvalidFormatException, IOException {
		if (workbook == null) {
			workbook = WorkbookFactory.create(is);
		}
		return workbook;
	}

	protected Sheet getSheet(String name) throws InvalidFormatException,
			IOException {
		Workbook workbook = getWorkbook();
		return workbook.getSheet(name);
	}

	public RangeIterator getRangeIterator() {
		return rangeIterator;
	}

	public void setRangeIterator(RangeIterator rangeIterator) {
		this.rangeIterator = rangeIterator;
	}

	private RangeIterator rangeIterator;

	private Point lastProcessedPosition = new Point(0,0);

	public Point getLastProcessedPosition() {
		return lastProcessedPosition;
	}

	public void saveLastProcessedPosition() {
		lastProcessedPosition.getX().setValue(positionPoint.getX().getValue());
		lastProcessedPosition.getY().setValue(positionPoint.getY().getValue());
	}

	private Point positionPoint = new Point();

	protected void resetPositionPoint() {
		positionPoint.getX().setValue(0);
		positionPoint.getY().setValue(0);
		saveLastProcessedPosition();
	}

	protected Point getPositionPoint() {
		return positionPoint;
	}

}
