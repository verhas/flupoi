package javax0.flupoi;

class Point {
	final private Coordinate x = new Coordinate();
	final private Coordinate y = new Coordinate();

	public Point() {
	}

	public Point(int x, int y) {
		getX().setValue(x);
		getY().setValue(y);
	}

	public Coordinate getX() {
		return x;
	}

	public Coordinate getY() {
		return y;
	}

	public String toString() {
		return "(" + x + "," + y + ")";
	}
}
