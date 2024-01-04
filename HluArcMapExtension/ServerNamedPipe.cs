// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2011 Hampshire Biodiversity Information Centre
// 
// This file is part of HLUTool.
// 
// HLUTool is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
// 
// HLUTool is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
// 
// You should have received a copy of the GNU General Public License
// along with HLUTool.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using AppModule.NamedPipes;
using HLU;
using HLU.GISApplication.ArcGIS;

namespace Server
{
    public sealed class ServerNamedPipe : IDisposable
    {
        #region Fields

        internal Thread PipeThread;
        internal ServerPipeConnection PipeConnection;
        internal bool Listen = true;
        internal DateTime LastAction;
        private bool disposed = false;
        private bool _sendResponse;
        private string _stringContinue = ArcMapApp.PipeStringContinue.ToString();

        #endregion

        private void PipeListener()
        {
            CheckIfDisposed();

            try
            {
                Listen = HluArcMapExtensionV4.PipeManager.Listen;
                HluArcMapExtensionV4.PipeData = new List<string>();
                StringBuilder sbRequest;

                bool continueString = false;

                while (Listen)
                {
                    LastAction = DateTime.Now;

                    string request = PipeConnection.Read();
                    while (!String.IsNullOrEmpty(request) && (request != "@"))
                    {
                        if (request == _stringContinue)
                        {
                            continueString = true;
                        }
                        else
                        {
                            if (continueString)
                            {
                                sbRequest = new StringBuilder(
                                    HluArcMapExtensionV4.PipeData[HluArcMapExtensionV4.PipeData.Count - 1]);
                                HluArcMapExtensionV4.PipeData[HluArcMapExtensionV4.PipeData.Count - 1] = 
                                    sbRequest.Append(request).ToString();
                                continueString = false;
                            }
                            else
                            {
                                HluArcMapExtensionV4.PipeData.Add(request);
                            }
                        }
                        request = PipeConnection.Read();
                    }

                    if ((HluArcMapExtensionV4.PipeData.Count > 0) && (request == "@"))
                    {
                        // wire event to be notified of outgoing data ready
                        HluArcMapExtensionV4.OutgoingDataReady += new EventHandler(HluArcMapExtensionV4_OutgoingDataReady);

                        // raise event in HluArcMapExtensionV4
                        HluArcMapExtensionV4.PipeManager.HandleRequest(String.Empty);

                        // unwire event
                        HluArcMapExtensionV4.OutgoingDataReady -= HluArcMapExtensionV4_OutgoingDataReady;
                        
                        // send response
                        foreach (string s in HluArcMapExtensionV4.PipeData)
                            PipeConnection.Write(s);

                        HluArcMapExtensionV4.PipeData.Clear();
                        PipeConnection.Write("@");
                    }

                    LastAction = DateTime.Now;
                    PipeConnection.Disconnect();
                    
                    if (Listen)
                    {
                        Connect();
                    }

                    HluArcMapExtensionV4.PipeManager.WakeUp();
                }
            }
            catch (System.Threading.ThreadAbortException ex) { }
            catch (System.Threading.ThreadStateException ex) { }
            catch (Exception ex)
            {
                // Log exception
            }
            finally
            {
                this.Close();
            }
        }

        void HluArcMapExtensionV4_OutgoingDataReady(object sender, EventArgs e)
        {
            _sendResponse = true;
        }

        internal void Connect()
        {
            CheckIfDisposed();
            PipeConnection.Connect();
        }
        
        internal void Close()
        {
            CheckIfDisposed();
            this.Listen = false;
            HluArcMapExtensionV4.PipeManager.RemoveServerChannel(this.PipeConnection.NativeHandle);
            this.Dispose();
        }

        internal void Start()
        {
            CheckIfDisposed();
            PipeThread.Start();
        }

        private void CheckIfDisposed()
        {
            if (this.disposed)
            {
                throw new ObjectDisposedException("ServerNamedPipe");
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (!this.disposed)
            {
                PipeConnection.Dispose();
                if (PipeThread != null)
                {
                    try
                    {
                        PipeThread.Abort();
                    }
                    catch (System.Threading.ThreadAbortException ex) { }
                    catch (System.Threading.ThreadStateException ex) { }
                    catch (Exception ex)
                    {
                        // Log exception
                    }
                }
            }
            disposed = true;
        }

        ~ServerNamedPipe()
        {
            Dispose(false);
        }

        internal ServerNamedPipe(string name, uint outBuffer, uint inBuffer, int maxReadBytes, bool secure)
        {
            PipeConnection = new ServerPipeConnection(name, outBuffer, inBuffer, maxReadBytes, secure);
            PipeThread = new Thread(new ThreadStart(PipeListener));
            PipeThread.IsBackground = true;
            PipeThread.Name = "Pipe Thread " + this.PipeConnection.NativeHandle.ToString();
            LastAction = DateTime.Now;
        }
    }
}