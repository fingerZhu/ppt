package util;

public enum FontStyle {

    FontSize("FontSize"), FontFamily("FontFamily"), FontColor("FontColor");

    private final String name;

    private FontStyle(String name) {
        this.name = name;
    }

    @Override
    public String toString() {
        return this.name;
    }
}
