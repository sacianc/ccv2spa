﻿// <copyright file="NotificationDataEntity.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator.Common.Repositories.NotificationData
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Azure.Cosmos.Table;
    using Newtonsoft.Json;

    /// <summary>
    /// Notification data entity class.
    /// </summary>
    public class NotificationDataEntity : TableEntity
    {
        /// <summary>
        /// Gets or sets Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets Title value.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the Image Link value.
        /// </summary>
        public string ImageLink { get; set; }

        /// <summary>
        /// Gets or sets the Summary value.
        /// </summary>
        public string Summary { get; set; }

        /// <summary>
        /// Gets or sets the Author value.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Gets or sets the Button Title value.
        /// </summary>
        public string ButtonTitle { get; set; }

        /// <summary>
        /// Gets or sets the Button Link value.
        /// </summary>
        public string ButtonLink { get; set; }

        /// <summary>
        /// Gets or sets the Secondary button Title value.
        /// </summary>
        public string ButtonTitle2 { get; set; }

        /// <summary>
        /// Gets or sets the Secondary button Link value.
        /// </summary>
        public string ButtonLink2 { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether scheduled or not.
        /// </summary>
        public bool IsScheduled { get; set; }

        /// <summary>
        /// Gets or sets the Schedule DateTime value.
        /// </summary>
        public DateTime ScheduleDate { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether recurrence message or not.
        /// </summary>
        public bool IsRecurrence { get; set; }

        /// <summary>
        /// Gets or sets the Repeats value (EveryWeekday/Daily/Weekly/Monthly/Yearly/Custom).
        /// </summary>
        public string Repeats { get; set; }

        /// <summary>
        /// Gets or sets the Repeat for value.
        /// </summary>
        public int RepeatFor { get; set; }

        /// <summary>
        /// Gets or sets the Repeat frequency value (Day/Week/Month).
        /// </summary>
        public string RepeatFrequency { get; set; }

        /// <summary>
        /// Gets or sets the Week Selection value (0/1/2/3/4/5/6).
        /// </summary>
        public string WeekSelection { get; set; }

        /// <summary>
        /// Gets or sets the repeat StartDate DateTime value.
        /// </summary>
        public DateTime RepeatStartDate { get; set; }

        /// <summary>
        /// Gets or sets the repeat EndDate DateTime value.
        /// </summary>
        public DateTime RepeatEndDate { get; set; }

        /// <summary>
        /// Gets or sets the CreatedBy value.
        /// </summary>
        public string CreatedBy { get; set; }

        /// <summary>
        /// Gets or sets the Created DateTime value.
        /// </summary>
        public DateTime CreatedDate { get; set; }

        /// <summary>
        /// Gets or sets the Sent DateTime value.
        /// </summary>
        public DateTime? SentDate { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the notification is sent out or not.
        /// </summary>
        public bool IsDraft { get; set; }

        /// <summary>
        /// Gets or sets TeamsInString value.
        /// This property helps to save the Teams data in Azure Table storage.
        /// Table Storage doesn't support array type of property directly.
        /// </summary>
        public string TeamsInString { get; set; }

        /// <summary>
        /// Gets or sets Teams audience collection.
        /// </summary>
        [IgnoreProperty]
        public IEnumerable<string> Teams
        {
            get
            {
                return JsonConvert.DeserializeObject<IEnumerable<string>>(this.TeamsInString);
            }

            set
            {
                this.TeamsInString = JsonConvert.SerializeObject(value);
            }
        }

        /// <summary>
        /// Gets or sets RostersInString value.
        /// This property helps to save the Rosters list in Table Storage.
        /// Table Storage doesn't support array type of property directly.
        /// </summary>
        public string RostersInString { get; set; }

        /// <summary>
        /// Gets or sets Rosters audience collection.
        /// </summary>
        [IgnoreProperty]
        public IEnumerable<string> Rosters
        {
            get
            {
                return JsonConvert.DeserializeObject<IEnumerable<string>>(this.RostersInString);
            }

            set
            {
                this.RostersInString = JsonConvert.SerializeObject(value);
            }
        }

        /// <summary>
        /// Gets or sets ADGroupsInString value.
        /// This property helps to save the AD Groups in Table Storage.
        /// Table Storage doesn't support array type of property directly.
        /// </summary>
        public string ADGroupsInString { get; set; }

        /// <summary>
        /// Gets or sets AD Groups collection.
        /// </summary>
        [IgnoreProperty]
        public IEnumerable<string> ADGroups
        {
            get
            {
                return JsonConvert.DeserializeObject<IEnumerable<string>>((!string.IsNullOrEmpty(this.ADGroupsInString)) ? this.ADGroupsInString : string.Empty);
            }

            set
            {
                this.ADGroupsInString = JsonConvert.SerializeObject(value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether a notification should be sent to all the users.
        /// </summary>
        public bool AllUsers { get; set; }

        /// <summary>
        /// Gets or sets message version number.
        /// </summary>
        public string MessageVersion { get; set; }

        /// <summary>
        /// Gets or sets the number of recipients who have received the notification successfully.
        /// </summary>
        public int Succeeded { get; set; }

        /// <summary>
        /// Gets or sets the number of recipients who failed in receiving the notification.
        /// </summary>
        public int Failed { get; set; }

        /// <summary>
        /// Gets or sets the number of recipients who were throttled out.
        /// </summary>
        public int Throttled { get; set; }

        /// <summary>
        /// Gets or sets the number of recipients acknowledged.
        /// </summary>
        public int MessageAcknowledged { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the sending process is completed or not.
        /// </summary>
        public bool IsCompleted { get; set; }

        /// <summary>
        /// Gets or sets the total number of expected messages to send.
        /// </summary>
        public int TotalMessageCount { get; set; }

        /// <summary>
        /// Gets or sets the sending started DateTime value.
        /// </summary>
        public DateTime? SendingStartedDate { get; set; }
    }
}