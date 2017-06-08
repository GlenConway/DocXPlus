namespace DocXPlus
{
    /// <summary>
    /// Defines spacing
    /// </summary>
    public class SpacingBetweenLines
    {
        private DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines spacing;

        /// <summary>
        ///
        /// </summary>
        /// <param name="spacing"></param>
        internal SpacingBetweenLines(DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines spacing)
        {
            this.spacing = spacing;
        }

        /// <summary>
        /// Spacing Below Paragraph.
        /// </summary>
        public string After
        {
            get
            {
                return spacing.After;
            }
            set
            {
                spacing.After = value;
            }
        }

        /// <summary>
        /// Automatically Determine Spacing Below Paragraph.
        /// </summary>
        public bool AfterAutoSpacing
        {
            get
            {
                return spacing.AfterAutoSpacing;
            }
            set
            {
                spacing.AfterAutoSpacing = value;
            }
        }

        /// <summary>
        /// Spacing Below Paragraph in Line Units.
        /// </summary>
        public int AfterLines
        {
            get
            {
                return spacing.AfterLines;
            }
            set
            {
                spacing.AfterLines = value;
            }
        }

        /// <summary>
        /// Spacing Above Paragraph.
        /// </summary>
        public string Before
        {
            get
            {
                return spacing.Before;
            }
            set
            {
                spacing.Before = value;
            }
        }

        /// <summary>
        /// Automatically Determine Spacing Above Paragraph.
        /// </summary>
        public bool BeforeAutoSpacing
        {
            get
            {
                return spacing.BeforeAutoSpacing;
            }
            set
            {
                spacing.BeforeAutoSpacing = value;
            }
        }

        /// <summary>
        /// Spacing Above Paragraph IN Line Units.
        /// </summary>
        public int BeforeLines
        {
            get
            {
                return spacing.BeforeLines;
            }
            set
            {
                spacing.BeforeLines = value;
            }
        }

        /// <summary>
        /// Spacing Between Lines in Paragraph.
        /// </summary>
        public string Line
        {
            get
            {
                return spacing.Line;
            }
            set
            {
                spacing.Line = value;
            }
        }

        /// <summary>
        /// Type of Spacing Between Lines.
        /// </summary>
        public LineSpacingRuleValue LineRule
        {
            get
            {
                return Convert.ToLineSpacingRuleValue(spacing.LineRule);
            }
            set
            {
                spacing.LineRule = Convert.ToLineSpacingRuleValues(value);
            }
        }
    }
}