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

	protected void setDirection(RangeDirection direction) {
		this.direction = direction;
	}

	protected Range() {
	}

}
