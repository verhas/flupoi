package javax0.flupoi;

import java.io.IOException;
import java.util.Collection;
import java.util.Iterator;
import java.util.LinkedList;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.eval.NotImplementedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;

class RangeIterator implements Iterable<Collection<Cell>>,
		Iterator<Collection<Cell>> {

	Processor processor;

	protected RangeIterator(final Processor processor) {
		this.processor = processor;
	}

	@Override
	public Iterator<Collection<Cell>> iterator() {
		final Point basePoint = processor.getProcessState()
				.getLastProcessedPosition();
		processor.getRange().absolutize(basePoint);
		calculateLimits();
		Point startPoint = processor.getRange().getStart();
		Point endPoint = processor.getRange().getEnd();
		final int a, b;
		switch (processor.getRange().getDirection()) {
		case RIGHT:
			a = startPoint.getX().getValue();
			b = endPoint.getX().getValue();
			setPosition(Math.min(a, b));
			break;
		case LEFT:
			a = startPoint.getX().getValue();
			b = endPoint.getX().getValue();
			setPosition(Math.max(a, b));
			break;
		case DOWN:
			a = startPoint.getY().getValue();
			b = endPoint.getY().getValue();
			setPosition(Math.min(a, b));
			break;
		case UP:
			a = startPoint.getY().getValue();
			b = endPoint.getY().getValue();
			setPosition(Math.max(a, b));
			break;
		}
		return this;
	}

	private Coordinate getPositionCoordinate() {
		final Point point = processor.getProcessState().getPositionPoint();
		final Coordinate cord;
		if (isVerticalRange()) {
			cord = point.getY();
		} else {
			cord = point.getX();
		}
		return cord;
	}

	private void stepPosition() {
		final Coordinate cord = getPositionCoordinate();
		processor.getProcessState().saveLastProcessedPosition();
		cord.setValue(cord.getValue() + (isForwardRange() ? 1 : -1));
	}

	private Collection<Cell> cells = null;

	private boolean isForwardRange() {
		RangeDirection direction = processor.getRange().getDirection();
		return direction == RangeDirection.RIGHT
				|| direction == RangeDirection.DOWN;
	}

	private boolean isVerticalRange() {
		return processor.getRange().isVertical();
	}

	private Sheet sheet = null;

	private Sheet getSheet() throws InvalidFormatException, IOException {
		if (sheet == null) {
			sheet = processor.getProcessState().getSheet(
					processor.getSheetName());
		}
		return sheet;
	}

	private int startInner = 0;
	private int endInner = 0;
	private int startOuter = 0;
	private int endOuter = 0;

	private void calculateLimits() {
		final Range range = processor.getRange();
		final Point startPoint = range.getStart();
		final Point endPoint = range.getEnd();
		final Coordinate startCord;
		final Coordinate endCord;
		if (isVerticalRange()) {
			startCord = startPoint.getX();
			endCord = endPoint.getX();
		} else {
			startCord = startPoint.getY();
			endCord = endPoint.getY();
		}
		final int a = startCord.getValue();
		final int b = endCord.getValue();
		startInner = Math.min(a, b);
		endInner = Math.max(a, b);

		if (isVerticalRange()) {
			startOuter = startPoint.getY().getValue();
			endOuter = endPoint.getY().getValue();
		} else {
			startOuter = startPoint.getX().getValue();
			endOuter = endPoint.getX().getValue();
		}

	}

	private void collectCells() {
		if (getPosition() >= 0) {
			cells = new LinkedList<Cell>();
			for (int i = startInner; i <= endInner; i++) {
				try {
					final int colNum = isVerticalRange() ? i : getPosition();
					final int rowNum = isVerticalRange() ? getPosition() : i;
					cells.add(getSheet().getRow(rowNum).getCell(colNum));
				} catch (NullPointerException | InvalidFormatException
						| IOException e) {
					cells = null;
					return;
				}
			}

		} else {
			cells = null;
		}
	}

	private boolean stopNow;

	private boolean isInInterval(int a, int b, int x) {
		if (a > b) {
			int swap = a;
			a = b;
			b = swap;
		}
		return a <= x && x <= b;
	}

	private boolean calculateStopOnRange() {
		return !isInInterval(startOuter, endOuter, getPosition());
	}

	private boolean calculateStopCondition() {
		return cells == null ? true : processor.getCondition().match(cells);
	}

	@Override
	public boolean hasNext() {
		if (processor.getCondition() == null) {
			stopNow = calculateStopOnRange();
		} else {
			if (cells == null) {
				collectCells();
				stopNow = calculateStopCondition();
			}
		}
		return !stopNow;
	}

	@Override
	public Collection<Cell> next() {
		if (cells == null) {
			collectCells();
		}
		final Collection<Cell> nextCells;
		if (hasNext()) {
			nextCells = cells;
			cells = null;
			stepPosition();
		} else {
			nextCells = null;
		}
		return nextCells;
	}

	@Override
	public void remove() {
		throw new NotImplementedException(
				"remove() is not implemented in RangeIterator");
	}

	public int getPosition() {
		return getPositionCoordinate().getValue();
	}

	protected void setPosition(final int position) {
		getPositionCoordinate().setValue(position);
	}

}
