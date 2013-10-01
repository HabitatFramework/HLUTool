// HLUTool is used to view and maintain habitat and land use GIS data.
// Copyright © 2013 Andy Foy
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

namespace AppModule.InterProcessComm
{
    /// <summary>
    /// Interface, which defines methods for a Channel Manager class.
    /// </summary>
    /// <remarks>
    /// A Channel Manager is responsible for creating and maintaining channels for inter-process communication. 
    /// The opened channels are meant to be reusable for performance optimization. Each channel needs to procees 
    /// requests by calling the <see cref="AppModule.InterProcessComm.IChannelManager.HandleRequest">HandleRequest</see> 
    /// method of the Channel Manager.
    /// </remarks>
    public interface IChannelManager
    {
        /// <summary>
        /// Initializes the Channel Manager.
        /// </summary>
        void Initialize(string pipeName, int maxReadBytes);

        /// <summary>
        /// Closes all opened channels and stops the Channel Manager.
        /// </summary>
        void Stop();
        
        /// <summary>
        /// Handles a request.
        /// </summary>
        /// <remarks>
        /// This method currently caters for text based requests. 
        /// XML strings can be used in case complex request structures are needed.
        /// </remarks>
        /// <param name="request">The incoming request.</param>
        /// <returns>The resulting response.</returns>
        string HandleRequest(string request);
        
        /// <summary>
        /// Indicates whether the Channel Manager is in listening mode.
        /// </summary>
        /// <remarks>
        /// This property is left public so that other classes, like a server channel 
        /// can start or stop listening based on the Channel Manager mode.
        /// </remarks>
        bool Listen { get; set; }
        
        /// <summary>
        /// Forces the Channel Manager to exit a sleeping mode and create a new channel.
        /// </summary>
        /// <remarks>
        /// Normally the Channel Manager will create a number of reusable channels, which will handle the incoming reqiests, and go into a sleeping mode. However if the request load is high, the Channel Manager needs to be asked to create additional channels.
        /// </remarks>
        void WakeUp();
        
        /// <summary>
        /// Removes an existing channel.
        /// </summary>
        /// <param name="param">A parameter identifying the channel.</param>
        void RemoveServerChannel(object param);

        /// <summary>
        /// Event raised when a batch of incoming data has been received.
        /// </summary>
        event EventHandler IncomingDataReady;
    }
}
