public static class ComClassWrapper
{
    public static dynamic ComObject(string progId)
    {
        var comObject = Type.GetTypeFromProgID(progId);
        return Activator.CreateInstance(comObject);
    }
}
