using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace DocXPlus
{
    /// <summary>
    /// Interface for the Container
    /// </summary>
    public interface IContainer
    {
        /// <summary>
        /// The width of the container
        /// </summary>
        UInt32Value AvailableWidth { get; }

        /// <summary>
        /// Adds an image to the container which can then be added to a paragraph
        /// </summary>
        /// <param name="data"></param>
        /// <param name="contentType"></param>
        /// <param name="width">The width of the image in English Metric Units (EMU)</param>
        /// <param name="height">The height of the image in English Metric Units (EMU)</param>
        /// <param name="name">The name of the image. Leave blank to generate a Guid</param>
        /// <returns></returns>
        Drawing AddImage(byte[] data, string contentType, Int64Value width, Int64Value height, string name);

        /// <summary>
        /// Adds an image to the container which can then be added to a paragraph
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="contentType"></param>
        /// <param name="width">The width of the image in English Metric Units (EMU)</param>
        /// <param name="height">The height of the image in English Metric Units (EMU)</param>
        /// <param name="name">The name of the image. Leave blank to generate a Guid</param>
        Drawing AddImage(Stream stream, string contentType, Int64Value width, Int64Value height, string name);

        /// <summary>
        /// Adds an image to the container which can then be added to a paragraph
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="x">The width of the image in English Metric Units (EMU)</param>
        /// <param name="y">The height of the image in English Metric Units (EMU)</param>
        /// <returns></returns>
        Drawing AddImage(string fileName, Int64Value x, Int64Value y);

        /// <summary>
        /// Adds a paragraph
        /// </summary>
        /// <returns></returns>
        Paragraph AddParagraph();

        /// <summary>
        /// Adds a paragraph with the supplied text
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        Paragraph AddParagraph(string text);

        /// <summary>
        /// Adds a paragraph with the supplied text and sets the alignment
        /// </summary>
        /// <param name="text"></param>
        /// <param name="alignment"></param>
        /// <returns></returns>
        Paragraph AddParagraph(string text, Align alignment);

        /// <summary>
        /// Adds a Table to the container with the specified number of columns
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <returns></returns>
        Table AddTable(int numberOfColumns);

        /// <summary>
        /// Adds a Table to the container with the specified number of columns using the percent widths
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <param name="percent"></param>
        /// <returns></returns>
        Table AddTable(int numberOfColumns, params int[] percent);

        /// <summary>
        /// Adds a Table to the container with the specified number of columns using the supplied widths
        /// </summary>
        /// <param name="numberOfColumns"></param>
        /// <param name="widths">The widths of the columns in Twips, CM or Inch. Leave blank for Twips, end with cm for centimetres or in for inches</param>
        /// <returns></returns>
        Table AddTable(int numberOfColumns, params string[] widths);
    }
}