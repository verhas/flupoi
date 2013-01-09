package javax0.flupoi;

import java.util.StringTokenizer;

public class Range {
	private String range;

	protected Range(final String range) {
		this.range = range;
		parse();
	}

	private void removeAllSpaces() {
		range = range.replaceAll("\\s+", "");
	}

	private void lowerCase() {
		range = range.toLowerCase();
	}

	private boolean isNxMyFormat(final String coordinate) {
		return coordinate.charAt(0) != '(';
	}

	private String rangeStart;
	private String rangeEnd;

	private void split() {
		final StringTokenizer st = new StringTokenizer(range, ":");
		rangeStart = st.nextToken();
		if (st.hasMoreTokens()) {
			rangeEnd = st.nextToken();
		}
	}

	private Point convertNumericFormat(final String coordinate) {
		throw new IllegalArgumentException("'" + coordinate
				+ "' is not a valid coordinate yet.");
	}

	private Point convertNxMyFormat(final String coordinate) {
		Point point = new Point();
		String cord = coordinate;
		int x = 0;
		if (cord.length() > 0 && cord.charAt(0) == '*') {
			cord = cord.substring(1);
		} else {
			while (cord.length() > 0 && Character.isAlphabetic(cord.charAt(0))) {
				x = 26 * x + (int) (cord.charAt(0) - 'a');
				cord = cord.substring(1);
			}
			point.setX(x);
		}
		int y = 0;
		if (cord.length() > 0 && cord.charAt(0) == '*') {
			cord = cord.substring(1);
		} else {
			while (cord.length() > 0 && Character.isDigit(cord.charAt(0))) {
				y = 10 * y + (int) (cord.charAt(0) - '0');
				cord = cord.substring(1);
			}
			point.setY(y-1);
		}
		if (cord.length() > 0) {
			throw new IllegalArgumentException("'" + coordinate
					+ "' is not a valid coordinate.");
		}
		return point;
	}

	private Point convert(final String coordinate) {
		final Point point;
		if (isNxMyFormat(coordinate)) {
			point = convertNxMyFormat(coordinate);
		} else {
			point = convertNumericFormat(coordinate);
		}
		return point;
	}

	private Point start;
	private Point end;

	public Point getStart() {
		return start;
	}

	public Point getEnd() {
		return end;
	}

	private RangeDirection direction;

	private void calculateDirection() {
		int joker = 0;
		if (!start.xDefined()) {
			direction = RangeDirection.LEFT;
			joker++;
		}
		if (!start.yDefined()) {
			direction = RangeDirection.UP;
			joker++;
		}
		if (!end.xDefined()) {
			direction = RangeDirection.RIGHT;
			joker++;
		}
		if (!end.yDefined()) {
			direction = RangeDirection.DOWN;
			joker++;
		}
		if (joker > 1) {
			throw new IllegalArgumentException("'" + range + "' is invalid.");
		}
	}

	public RangeDirection getDirection() {
		return direction;
	}

	public void setDirection(RangeDirection direction) {
		this.direction = direction;
	}

	/**
	 * Nx:My or (x,y):(x,y)
	 */
	private void parse() {
		removeAllSpaces();
		lowerCase();
		split();
		start = convert(rangeStart);
		if (rangeEnd != null) {
			end = convert(rangeEnd);
		} else {
			end = start;
		}
		calculateDirection();
	}

}
