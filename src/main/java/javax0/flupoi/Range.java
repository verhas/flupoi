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
		if( direction == null ){
			direction = RangeDirection.DOWN;
		}
		return direction;
	}

	protected boolean isVertical() {
		return getDirection() == RangeDirection.UP
				|| getDirection() == RangeDirection.DOWN;
	}

	protected void setDirection(RangeDirection direction) {
		this.direction = direction;
	}

	protected Range() {
	}

	private void absolutize(Coordinate coord, Coordinate baseCoordinate, int correction) {
		if (coord.isDefined() && coord.isRelative()) {
			coord.setValue(coord.getValue() * coord.getRelativity()
					+ baseCoordinate.getValue()+correction);
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
		absolutize(start.getX(), basePoint.getX(),0);
		absolutize(start.getY(), basePoint.getY(),0);
		absolutize(end.getX(), start.getX(),-1);
		absolutize(end.getY(), start.getY(),-1);
	}
}
