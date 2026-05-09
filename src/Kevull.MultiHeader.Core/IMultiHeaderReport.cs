namespace Kevull.MultiHeader.Core
{
    /// <summary>
    /// Defines the operations required to configure, generate, and save a multi-header report.
    /// </summary>
    /// <typeparam name="T">The type of the data items used to populate the report.</typeparam>
    public interface IMultiHeaderReport<T>
    {
        /// <summary>
        /// Configures the report using the specified configuration action.
        /// </summary>
        /// <param name="options">An action that configures the report columns and options.</param>
        /// <returns>The current report instance to allow fluent configuration.</returns>
        IMultiHeaderReport<T> Configure(Action<IConfigurationBuilder<T>> options);

        /// <summary>
        /// Generates the report content from the provided data.
        /// </summary>
        /// <param name="data">The sequence of data items to include in the report.</param>
        void GenerateReport(IEnumerable<T> data);

        /// <summary>
        /// Saves the generated report to the specified file.
        /// </summary>
        /// <param name="fileName">The destination file name or path.</param>
        void Save(string fileName);
    }
}