using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.AddIn
{
    [Serializable]
    public delegate void RibbonControlClickEventHandler(RibbonControl control);

    [Serializable]
    public delegate void RibbonControlHelpEventHandler(RibbonControl control, IntPtr hWndFrame, int helpCommandID);

    /// <summary>
    /// Abstract base class for all ribbon controls.
    /// </summary>
    public abstract class RibbonControl
    {
        /// <summary>
        /// Raised when a user clicks the control.
        /// </summary>
        public event RibbonControlClickEventHandler Click;

        /// <summary>
        /// Raised when user requests help for the control.
        /// </summary>
        public event RibbonControlHelpEventHandler Help;

        private bool _checked = false;
        private int _commandId = -1;
        private string _commandName;
        private bool _enabled = true;
        private RibbonGroup _group;
        private int _imageId = -1;
        private string _label;
        private string _macro;
        private string _macroParameters;
        private string _screenTip;
        private bool _showImage = true;
        private bool _showLabel = true;
        private int _solidEdgeCommandId = -1;
        private string _superTip;

        internal RibbonControl(int commandId)
        {
            _commandId = commandId;
        }

        public bool UseDotMark;

        #region Properties

        /// <summary>
        /// Gets or set a value indicating whether the control is in the checked state.
        /// </summary>
        public bool Checked { get { return _checked; } set { _checked = value; } }

        /// <summary>
        /// Gets the command id of the control.
        /// </summary>
        /// <remarks>
        /// This is the command id used by OnCommand SolidEdgeFramework.AddInEvents.
        /// </remarks>
        public int CommandId { get { return _commandId; } }

        /// <summary>
        /// Returns the generated command name used when calling SetAddInInfo().
        /// </summary>
        public string CommandName { get { return _commandName; } internal set { _commandName = value; } }

        /// <summary>
        /// Gets or set a value indicating whether the control is enabled.
        /// </summary>
        public bool Enabled { get { return _enabled; } set { _enabled = value; } }

        /// <summary>
        /// Gets ribbon group that the control is assigned to.
        /// </summary>
        public RibbonGroup Group
        {
            get { return _group; }
            internal set { _group = value; }
        }

        /// <summary>
        /// Gets or set a value referencing an image embedded into the assembly as a native resource using the NativeResourceAttribute.
        /// </summary>
        /// <remarks>
        /// Changing this value after the ribbon has been initialized has no impact.
        /// </remarks>
        public int ImageId { get { return _imageId; } set { _imageId = value; } }

        /// <summary>
        /// Gets or set a value indicating the label (caption) of the control.
        /// </summary>
        public string Label { get { return _label; } set { _label = value; } }

        /// <summary>
        /// Gets or set a value indicating a macro (.exe) assigned to the control.
        /// </summary>
        public string Macro { get { return _macro; } set { _macro = value; } }

        /// <summary>
        /// Gets or set a value indicating the macro (.exe) parameters assigned to the control.
        /// </summary>
        public string MacroParameters { get { return _macroParameters; } set { _macroParameters = value; } }

        /// <summary>
        /// Gets or set a value indicating screentip of the control.
        /// </summary>
        /// <remarks>
        /// Changing this value after the ribbon has been initialized has no impact.
        /// </remarks>
        public string ScreenTip { get { return _screenTip; } set { _screenTip = value; } }

        /// <summary>
        /// Gets or set a value indicating whether to show the image assigned to the control.
        /// </summary>
        public bool ShowImage { get { return _showImage; } set { _showImage = value; } }

        /// <summary>
        /// Gets or set a value indicating whether to show the label (caption) assigned to the control.
        /// </summary>
        public bool ShowLabel { get { return _showLabel; } set { _showLabel = value; } }

        /// <summary>
        /// Gets the Solid Edge assigned runtime command id of the control.
        /// </summary>
        /// <remarks>
        /// This is the command id used by the BeforeCommand and AfterCommand SolidEdgeFramework.ApplicationEvents. 
        /// You also use this command id when calling SolidEdgeFramework.Application.StartCommand().
        /// </remarks>
        public int SolidEdgeCommandId
        {
            get { return _solidEdgeCommandId; }
            internal set { _solidEdgeCommandId = value; }
        }

        internal abstract SolidEdgeFramework.SeButtonStyle Style { get; }

        /// <summary>
        /// Gets or set a value indicating the supertip of the control.
        /// </summary>
        /// <remarks>
        /// Changing this value after the ribbon has been initialized has no impact.
        /// </remarks>
        public string SuperTip { get { return _superTip; } set { _superTip = value; } }

        #endregion

        #region TryParse helpers

        internal void TryParseEnabled(string enabled)
        {
            int iEnabled = 1;

            if (bool.TryParse(enabled, out _enabled))
            {
                // Case: true\false
            }
            else if (int.TryParse(enabled, out iEnabled))
            {
                // Case 1\0
                _enabled = iEnabled == 1 ? true : false;
            }
        }

        internal void TryParseImageId(string imageId)
        {
            if (int.TryParse(imageId, out _imageId) == false) { _imageId = -1; }
        }

        internal void TryParseShowImage(string showImage)
        {
            if (bool.TryParse(showImage, out _showImage) == false) { _showImage = true; }
        }

        internal void TryParseShowLabel(string showLabel)
        {
            if (bool.TryParse(showLabel, out _showLabel) == false) { _showLabel = false; }
        }

        #endregion

        #region Event invokers

        internal virtual void DoClick()
        {
            if (Click != null)
            {
                Click(this);
            }
        }

        internal void DoHelp(IntPtr hWndFrame, int helpCommandID)
        {
            if (Help != null)
            {
                Help(this, hWndFrame, helpCommandID);
            }
        }

        #endregion

    //    /// <summary>
    //    /// Properly formats the CommandName to be used in the CommandNames array of SetAddInInfoEx().
    //    /// </summary>
    //    internal string ToCommandName()
    //    {
    //        StringBuilder sb = new StringBuilder();

    //        var guid = Guid.NewGuid();
    //        var prefix = String.Format("{0}_{1}", guid.ToString().Substring(0, 5), CommandId);

    //        sb.AppendFormat("{0}\n{1}\n{2}\n{3}", prefix, Label, SuperTip, ScreenTip);

    //        // Append macro info if provided.
    //        if (!String.IsNullOrEmpty(Macro))
    //        {
    //            sb.AppendFormat("\n{0}", Macro);

    //            if (!String.IsNullOrEmpty(MacroParameters))
    //            {
    //                sb.AppendFormat("\n{0}", MacroParameters);
    //            }
    //        }

    //        return sb.ToString();
    //    }
    }

    /// <summary>
    /// Ribbon button control class.
    /// </summary>
    public class RibbonButton : RibbonControl
    {
        private RibbonButtonSize _size = RibbonButtonSize.Normal;

        internal RibbonButton(int id)
            : base(id)
        {
        }

        internal void TryParseSize(string size)
        {
            if (String.IsNullOrWhiteSpace(size) == false)
            {
                _size = size.Equals("large", StringComparison.OrdinalIgnoreCase) == true ? RibbonButtonSize.Large : RibbonButtonSize.Normal;
            }
        }

        /// <summary>
        /// Gets or sets the size of the control.
        /// </summary>
        public RibbonButtonSize Size
        {
            get { return _size; }
            set { _size = value; }
        }

        /// <summary>
        /// Returns the control style.
        /// </summary>
        internal override SolidEdgeFramework.SeButtonStyle Style
        {
            get
            {
                var style = SolidEdgeFramework.SeButtonStyle.seButtonAutomatic;

                if (Size == RibbonButtonSize.Large)
                {
                    style = SolidEdgeFramework.SeButtonStyle.seButtonIconAndCaptionBelow;
                }
                else
                {
                    if (ShowImage && ShowLabel)
                    {
                        style = SolidEdgeFramework.SeButtonStyle.seButtonIconAndCaption;
                    }
                    else if (ShowImage)
                    {
                        style = SolidEdgeFramework.SeButtonStyle.seButtonIcon;
                    }
                    else if (ShowLabel)
                    {
                        style = SolidEdgeFramework.SeButtonStyle.seButtonCaption;
                    }
                    else
                    {
                        style = SolidEdgeFramework.SeButtonStyle.seButtonAutomatic;
                    }
                }

                return style;
            }

        }
    }

    /// <summary>
    /// Ribbon checkbox control class.
    /// </summary>
    public class RibbonCheckBox : RibbonControl
    {
        internal RibbonCheckBox(int id)
            : base(id)
        {
        }

        internal override void DoClick()
        {
            // Toggle the check state before raising the event.
            this.Checked = !this.Checked;

            base.DoClick();
        }

        internal override SolidEdgeFramework.SeButtonStyle Style
        {
            get
            {
                var style = SolidEdgeFramework.SeButtonStyle.seCheckButton;

                if (this.ImageId >= 0)
                {
                    style = ShowImage == true ? SolidEdgeFramework.SeButtonStyle.seCheckButtonAndIcon : SolidEdgeFramework.SeButtonStyle.seCheckButton;
                }
                
                return style;
            }
        }
    }

    /// <summary>
    /// Ribbon radio button control class.
    /// </summary>
    public class RibbonRadioButton : RibbonControl
    {
        internal RibbonRadioButton(int id)
            : base(id)
        {
        }

        internal override void DoClick()
        {
            // Toggle the check state before raising the event.
            this.Checked = !this.Checked;

            base.DoClick();
        }

        internal override SolidEdgeFramework.SeButtonStyle Style { get { return SolidEdgeFramework.SeButtonStyle.seRadioButton; } }
    }

    /// <summary>
    /// Ribbon button sizes.
    /// </summary>
    public enum RibbonButtonSize
    {
        /// <summary>
        /// Standard size button.
        /// </summary>
        Normal,

        /// <summary>
        /// Large size button.
        /// </summary>
        Large
    }
}
