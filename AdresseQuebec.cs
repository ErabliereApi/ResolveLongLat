namespace ResolveLongLat;

public class AdressesQuebec 
{
    public SpatialReference? spatialReference { get; set; }

    public Candidate[]? candidates { get; set; }
}

public class SpatialReference
{
    public int wkid { get; set; }
    public int latestWkid { get; set; }
}

public class Candidate
{
    public string? address { get; set; }
    public Location? location { get; set; }
}

public class Location
{
    public double x { get; set; }
    public double y { get; set; }
}