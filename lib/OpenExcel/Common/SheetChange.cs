﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenExcel.Common
{
    /// <summary>
    /// Details of sheet change, e.g. inserted rows.
    /// </summary>
    public class SheetChange
    {
        /// <summary>
        /// Sheet where change occurred.
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// Starting row of change. Zero if no change in rows occurred.
        /// </summary>
        public uint RowStart { get; set; }

        /// <summary>
        /// Starting column of change. Zero if no change in columns occurred.
        /// </summary>
        public uint ColumnStart { get; set; }

        /// <summary>
        /// Rows inserted(+) or deleted (-).
        /// </summary>
        public int RowDelta { get; set; }

        /// <summary>
        /// Columns inserted(+) or deleted (-).
        /// </summary>
        public int ColumnDelta { get; set; }
    }
}
