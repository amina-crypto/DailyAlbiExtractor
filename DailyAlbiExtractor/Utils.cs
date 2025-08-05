using System;

namespace DailyAlbiExtractor
{
    public static class Utils
    {
        public static bool IsNumericType(Type type)
        {
            return type == typeof(int) || type == typeof(long) || type == typeof(double) || type == typeof(float) || type == typeof(decimal);
        }
    }
}