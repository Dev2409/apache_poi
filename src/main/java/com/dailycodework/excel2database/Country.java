package com.dailycodework.excel2database;

public enum Country {
    UNITED_STATES("United States"),
    GREAT_BRITAIN("Great Britain"),
    FRANCE("France");

    private final String displayName;

    Country(String displayName) {
        this.displayName = displayName;
    }

    public String getDisplayName() {
        return displayName;
    }

    public static Country fromDisplayName(String displayName) {
        for (Country country : Country.values()) {
            if (country.getDisplayName().equalsIgnoreCase(displayName)) {
                return country;
            }
        }
        throw new IllegalArgumentException("No enum constant for display name: " + displayName);
    }
}
