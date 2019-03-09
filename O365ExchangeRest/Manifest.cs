//-----------------------------------------------------------------------
// <copyright file="Manifest.cs" company="Microsoft Corporation">
//     Copyright (c) Microsoft Corporation. All rights reserved.
// </copyright>
// <author>Chris Stepanek (cstep)</author>
//-----------------------------------------------------------------------
namespace O365ExchangeRest
{
    using System;
    using System.Diagnostics.CodeAnalysis;

    [SuppressMessage("StyleCop.CSharp.NamingRules", "SA1300:ElementMustBeginWithUpperCaseLetter", Justification = "Manifest uses lower case property names.")]

    /// <summary>
    /// Defines the <see cref="Manifest"/> class.
    /// </summary>
    /// <remarks>Used for JSON serialization.</remarks>
    internal class Manifest
    {
        /// <summary>
        /// Gets or sets the keyID value.
        /// </summary>
        public Guid keyId { get; set; }

        /// <summary>
        /// Gets or sets the certificate raw data value.
        /// </summary>
        public string value { get; set; }

        /// <summary>
        /// Gets or sets the type value.
        /// </summary>
        public string type { get; set; }

        /// <summary>
        /// Gets or sets the usage value.
        /// </summary>
        public string usage { get; set; }

        /// <summary>
        /// Gets or sets the certificate hash value.
        /// </summary>
        public string customKeyIdentifier { get; set; }
    }
}
