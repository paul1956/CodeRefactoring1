﻿// Copyright (c) Microsoft.  All Rights Reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

using System.Collections.Immutable;

namespace Utilities
{
    public static class CodeRefactoringHash
    {
        /// <summary>
        /// This is how VB Anonymous Types combine hash values for fields.
        /// </summary>
        public static int Combine(int newKey, int currentKey)
        {
            return unchecked((currentKey * (int)0xA5555529) + newKey);
        }

        public static int Combine(bool newKeyPart, int currentKey)
        {
            return Combine(currentKey, newKeyPart ? 1 : 0);
        }

        /// <summary>
        /// This is how VB Anonymous Types combine hash values for fields.
        /// PERF: Do not use with enum types because that involves multiple
        /// unnecessary boxing operations.  Unfortunately, we can't constrain
        /// T to "non-enum", so we'll use a more restrictive constraint.
        /// </summary>
        public static int Combine<T>(T newKeyPart, int currentKey) where T : class
        {
            int hash = unchecked(currentKey * (int)0xA5555529);

            if (newKeyPart != null)
            {
                return unchecked(hash + newKeyPart.GetHashCode());
            }

            return hash;
        }

        public static int CombineValues<T>(ImmutableArray<T> values, int maxItemsToHash = int.MaxValue)
        {
            if (values.IsDefaultOrEmpty)
            {
                return 0;
            }

            var hashCode = 0;
            var count = 0;
            foreach (T value in values)
            {
                if (count++ >= maxItemsToHash)
                {
                    break;
                }

                // Should end up with a constrained virtual call to object.GetHashCode (i.e. avoid boxing where possible).
#pragma warning disable RECS0017 // Possible compare of value type with 'null'
                if (value != null)
#pragma warning restore RECS0017 // Possible compare of value type with 'null'
                {
                    hashCode = CodeRefactoringHash.Combine(value.GetHashCode(), hashCode);
                }
            }

            return hashCode;
        }

        /// <summary>
        /// The offset bias value used in the FNV-1a algorithm
        /// See http://en.wikipedia.org/wiki/Fowler%E2%80%93Noll%E2%80%93Vo_hash_function
        /// </summary>
        public const int FnvOffsetBias = unchecked((int)2166136261);

        /// <summary>
        /// The generative factor used in the FNV-1a algorithm
        /// See http://en.wikipedia.org/wiki/Fowler%E2%80%93Noll%E2%80%93Vo_hash_function
        /// </summary>
        public const int FnvPrime = 16777619;
    }
}
