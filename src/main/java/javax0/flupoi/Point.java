package javax0.flupoi;

class Point {
	private int x;
	private boolean xDef = false;
	private int y;
	private boolean yDef = false;

	public int getX() {
		return x;
	}

	public void setX(int x) {
		this.x = x;
		xDef = true;
	}

	public int getY() {
		return y;
	}

	public void setY(int y) {
		this.y = y;
		yDef = true;
	}

	public boolean xDefined() {
		return xDef;
	}

	public boolean yDefined() {
		return yDef;
	}
}
