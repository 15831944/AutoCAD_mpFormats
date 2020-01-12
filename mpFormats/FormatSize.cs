namespace mpFormats
{
    /// <summary>
    /// Format size
    /// </summary>
    public class FormatSize
    {
        /// <summary>
        /// Main constructor
        /// </summary>
        /// <param name="width">Width value</param>
        /// <param name="height">Height value</param>
        public FormatSize(double width, double height)
        {
            Width = width;
            Height = height;
        }

        /// <summary>
        /// Width
        /// </summary>
        public double Width { get; }

        /// <summary>
        /// Height
        /// </summary>
        public double Height { get; }
    }
}
