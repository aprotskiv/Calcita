/*****************************************************
 * 
 * ReoGrid - .NET Spreadsheet Control
 * 
 * https://reogrid.net/
 *
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR
 * PURPOSE.
 *
 * Author: Jingwood <jingwood at unvell.com>
 *
 * Copyright (c) 2012-2025 Jingwood <jingwood at unvell.com>
 * Copyright (c) 2012-2025 UNVELL Inc. All rights reserved.
 * 
 ****************************************************************************/namespace Calcita.Actions
{
	/// <summary>
	/// Action to remove data from specified range.
	/// </summary>
	public class RemoveRangeDataAction : WorksheetReusableAction
	{
		private object[,] backupData;

        /// <summary>
        /// Pressed keys that triggered <see cref="RemoveRangeDataAction"/>
        /// </summary>
        /// <remarks>
        /// Allows to customize <see cref="RemoveRangeDataAction"/>'s behavior
        /// </remarks>
        public Interaction.KeyCode? KeyData { get; }

        /// <summary>
        /// Create action to remove data from specified range.
        /// </summary>
        /// <param name="range">data from cells in this range will be removed.</param>
        /// <param name="keyData">pressed keys that triggered <see cref="RemoveRangeDataAction"/></param>
        public RemoveRangeDataAction(RangePosition range, Interaction.KeyCode? keyData = null)
            : base(range)
		{
            this.KeyData = keyData;
        }

		/// <summary>
		/// Create a copy from this action in order to apply the operation to another range.
		/// </summary>
		/// <param name="range">New range where this operation will be appiled to.</param>
		/// <returns>New action instance copied from this action.</returns>
		public override WorksheetReusableAction Clone(RangePosition range)
		{
            return new RemoveRangeDataAction(range, this.KeyData);
        }

		/// <summary>
		/// Do action to remove data from specified range.
		/// </summary>
		public override void Do()
		{
			this.backupData = Worksheet.GetRangeData(base.Range);
			this.Worksheet.DeleteRangeData(this.Range, true);
		}

		/// <summary>
		/// Undo action to restore removed data.
		/// </summary>
		public override void Undo()
		{
			this.Worksheet.SetRangeData(this.Range, this.backupData);
		}

		/// <summary>
		/// Get friendly name of this action.
		/// </summary>
		/// <returns>friendly name of this action.</returns>
		public override string GetName()
		{
			return "Remove Cells Data";
		}
	}
}


