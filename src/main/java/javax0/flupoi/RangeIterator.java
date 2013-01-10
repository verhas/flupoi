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
		switch (processor.getRange().getDirection()) {
		case RIGHT:
			setPosition(processor.getRange().getStart().getX().getValue());
			break;
		case LEFT:
			setPosition(processor.getRange().getEnd().getX().getValue());
			break;
		case DOWN:
			setPosition(processor.getRange().getStart().getY().getValue());
			break;
		case UP:
			setPosition(processor.getRange().getEnd().getY().getValue());
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

	private boolean stillInTheRow(int i, int end) {
		return isForwardRange() ? i <= end : i >= end;
	}

	private int stepIncrement() {
		return isForwardRange() ? +1 : -1;
	}

	private void collectCells() {
		if (getPosition() >= 0) {
			cells = new LinkedList<Cell>();
			int start = 0;
			int end = 0;
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
			start = startCord.getValue();
			end = endCord.getValue();
			for (int i = start; stillInTheRow(i, end); i += stepIncrement()) {
				try {
					if (sheet == null) {
						sheet = processor.getProcessState().getSheet(
								processor.getSheetName());
					}
					final int colNum = isVerticalRange() ? i : getPosition();
					final int rowNum = isVerticalRange() ? getPosition() : i;
					cells.add(sheet.getRow(rowNum).getCell(colNum));
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

	private boolean calculateStopOnRange() {
		final Range range = processor.getRange();
		final Point startPoint = range.getStart();
		final Point endPoint = range.getEnd();
		final int start, end;
		if (isVerticalRange()) {
			start = startPoint.getY().getValue();
			end = endPoint.getY().getValue();
		} else {
			start = startPoint.getX().getValue();
			end = endPoint.getX().getValue();
		}
		return getPosition() > end || getPosition() < start;
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
