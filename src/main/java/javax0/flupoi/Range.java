package javax0.flupoi;

class Range {
	protected Point start;
	protected Point end;

	protected void setStart(Point start) {
		this.start = start;
	}

	protected void setEnd(Point end) {
		this.end = end;
	}

	protected Point getStart() {
		return start;
	}

	protected Point getEnd() {
		return end;
	}

	protected RangeDirection direction;

	protected RangeDirection getDirection() {
		return direction;
	}

	protected boolean isVertical() {
		return direction == RangeDirection.UP
				|| direction == RangeDirection.DOWN;
	}

	protected void setDirection(RangeDirection direction) {
		this.direction = direction;
	}

	protected Range() {
	}

	private void absolutize(Coordinate coord, Coordinate baseCoordinate) {
		if (coord.isDefined() && coord.isRelative()) {
			coord.setValue(coord.getValue() * coord.getRelativity()
					+ baseCoordinate.getValue());
			coord.resetRelativity();
		}
	}

	/**
	 * Make all coordinates of the range absolute based on the point.
	 * 
	 * @param basePoint
	 *            the point to which the relative coordinates are to be
	 *            calculated. Note that both coordinates of the base point have
	 *            to be absolute.
	 */
	protected void absolutize(Point basePoint) {
		absolutize(start.getX(), basePoint.getX());
		absolutize(start.getY(), basePoint.getY());
		absolutize(end.getX(), start.getX());
		absolutize(end.getY(), start.getY());
	}
}
