﻿using Jpp.AddIn.MailAssistant.Abstracts;
using Microsoft.AppCenter.Crashes;
using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Jpp.AddIn.MailAssistant.Wrappers
{
    internal class ReportItemWrapper : IMoveable
    {
        private readonly Outlook.ReportItem _innerObject;

        public Type InnerObjectType => _innerObject.GetType();
        public string Id => _innerObject.PropertyAccessor.GetProperty(Constants.PR_INTERNET_MESSAGE_ID) as string;
        public string RestrictCriteria
        {
            get
            {
                var subject = _innerObject.Subject;

                return $"[Subject] = '{subject}'";
            }
        }
        public string Description => $"{_innerObject.CreationTime} | {_innerObject.Subject}";
        public string Folder => ((Outlook.Folder)_innerObject.Parent).Name;
        public int Size => _innerObject.Size;
        
        public ReportItemWrapper(Outlook.ReportItem item)
        {
            _innerObject = item ?? throw new ArgumentNullException(nameof(item));
        }

        public bool Equals(IMoveable other)
        {
            if (other == null) return false;
            if (other.InnerObjectType != InnerObjectType) return false;

            return other.Id == Id || other.Description == Description;
        }

        bool IMoveable.Move(Outlook.Folder folder)
        {
            Outlook.ReportItem moved = null;
            Outlook.Folder parent = null;

            try
            {
                moved = _innerObject.Move(folder);
                parent = moved.Parent;

                return parent.Name == folder.Name;
            }
            catch (Exception e)
            {
                Crashes.TrackError(e);
                return false;
            }
            finally
            {
                if (moved != null) Marshal.ReleaseComObject(moved);
                if (parent != null) Marshal.ReleaseComObject(parent);
            }
        }

        #region IDisposable Support
        private bool _disposedValue;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposedValue) return;

            if (disposing) { } // TODO: dispose managed objects.

            Marshal.ReleaseComObject(_innerObject);

            _disposedValue = true;
        }

        ~ReportItemWrapper()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
